Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Windows.Forms
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application

Public Class SetBlocksAlignments
    '=============================================================================================
    'разбить трассу блоками через промежуток
    <CommandMethod("SetBlocksAlignments")>
    Public Sub SetBlocksAlignments()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ссылка на чертеж .dwg
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
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

        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Try
            Using tr
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '=====================================================================================================
                'выбираем трассу
                Dim Align As Autodesk.Civil.DatabaseServices.Alignment = Nothing
                Dim ColAlign As ObjectIdCollection = CivilDoc.GetAlignmentIds
                If ColAlign.Count = 0 Then
                    MsgBox("На чертеже нет объектов типа ТРАССА. Макро прервано.")
                    Exit Sub
                ElseIf ColAlign.Count = 1 Then
                    Align = tr.GetObject(ColAlign.Item(0), OpenMode.ForRead)
                Else
                    '=============================================================
                    'выбираем трассу
                    Dim ppr1 As PromptEntityResult = ed.GetEntity("Выберите трассу: ")
                    Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ppr1.ObjectId, OpenMode.ForRead)
                    If TypeOf ent1 Is Autodesk.Civil.DatabaseServices.Alignment Then
                        Align = ent1
                    End If
                End If
                If IsNothing(Align) = True Then
                    MsgBox("Ошибка выбора ТРАССЫ. Макро прервано.")
                    Exit Sub
                End If
                '==========================================================================================================
                'выбираем образец блока
                Dim ArrayPropBlk As String(,) = Nothing
                Dim blkName As String = ""
                Dim sc As Scale3d = New Scale3d(1, 1, 1)
                Dim rot As Double = 0
                Dim Layerblk As String = "0"
                Dim per As PromptEntityResult = ed.GetEntity("Выберите образец блока: ")
                If per.Status = PromptStatus.Cancel Then Exit Sub
                'выбор полигона которому принадлежит этот островок
                Dim obj As Autodesk.AutoCAD.DatabaseServices.DBObject = tr.GetObject(per.ObjectId, OpenMode.ForRead) 'ссылка на объект
                If TypeOf obj Is BlockReference Then
                    Dim Blk As BlockReference = obj
                    sc = Blk.ScaleFactors
                    Layerblk = Blk.Layer
                    rot = Blk.Rotation
                    If Blk.IsDynamicBlock = True Then
                        Dim blockRef As BlockReference = tr.GetObject(per.ObjectId, OpenMode.ForRead)
                        Dim blockDyn As BlockTableRecord = tr.GetObject(blockRef.DynamicBlockTableRecord, OpenMode.ForRead)
                        blkName = blockDyn.Name
                    Else
                        blkName = Blk.Name
                    End If
                    ReDim Preserve ArrayPropBlk(1, 2)
                    ArrayPropBlk(0, 0) = "Name"
                    ArrayPropBlk(1, 0) = blkName
                    ArrayPropBlk(0, 1) = "Layer"
                    ArrayPropBlk(1, 1) = Blk.Layer
                    ArrayPropBlk(0, 2) = "Color"
                    ArrayPropBlk(1, 2) = Blk.ColorIndex
                Else
                    MsgBox("Выбранный примитив не является блоком. Макро прервано.")
                    Exit Sub
                End If
                If blkName Like "" Then
                    MsgBox("Выбранный примитив не является блоком. Макро прервано.")
                    Exit Sub
                End If
                '==============================================================================================================
                'выбираем интервал
                Dim pStrOpts As PromptDoubleOptions = New PromptDoubleOptions(vbLf & "Укажите шаг разбивки: ")
                pStrOpts.AllowZero = False
                pStrOpts.AllowNone = False
                Dim pStrRes As PromptDoubleResult = acDoc.Editor.GetDouble(pStrOpts)
                Dim interval As Double = Math.Abs(pStrRes.Value)

                '==============================================================================================================
                'выбираем смещение
                Dim pStrOpts1 As PromptDoubleOptions = New PromptDoubleOptions(vbLf & "Укажите смещение: ")
                pStrOpts1.DefaultValue = 0
                Dim pStrRes1 As PromptDoubleResult = acDoc.Editor.GetDouble(pStrOpts1)
                Dim UserOffset As Double = pStrRes1.Value

                Do
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'указываем сегмент трассы
                    Dim SelectLine As PromptEntityResult = ed.GetEntity("Выберите сегмент трассы: ")
                    If SelectLine.Status = PromptStatus.Cancel Then Exit Do
                    'выбор полигона которому принадлежит этот островок
                    Dim PicPoint As Point3d = SelectLine.PickedPoint
                    '============================================================================================
                    'находим пикет
                    Dim PK As Double = -1
                    Dim offsetPk As Double = -1
                    Try
                        Align.StationOffset(PicPoint.X, PicPoint.Y, PK, offsetPk)
                    Catch ex As Autodesk.Civil.PointNotOnEntityException
                        MsgBox("Не удалось определить пикет сегмента")
                    End Try
                    'перебираем все элементы трассы для их сортировки в процессе возрастания
                    For Each myAe As AlignmentEntity In Align.Entities
                        Dim AlignEnt As AlignmentEntityType = myAe.EntityType
                        Dim stLine As Double = 0
                        Dim enLine As Double = 0
                        Select Case AlignEnt
                            Case AlignmentEntityType.Line 'прямой участок
                                Dim AlignLine As AlignmentLine = myAe
                                stLine = AlignLine.StartStation
                                enLine = AlignLine.EndStation
                            Case AlignmentEntityType.Arc
                                Dim AlignArc As AlignmentArc = myAe
                                stLine = AlignArc.StartStation
                                enLine = AlignArc.EndStation
                            Case AlignmentEntityType.Spiral
                                If TypeOf myAe Is AlignmentSTS Then
                                    Dim AlignSTS As AlignmentSTS = myAe 'составная спираль спираль дуга спираль
                                    stLine = AlignSTS.StartStation
                                    enLine = AlignSTS.EndStation
                                    '==================================================================================================================================
                                ElseIf TypeOf myAe Is AlignmentSCS Then
                                    Dim AlignSCS As AlignmentSCS = myAe 'составная спираль спираль прямая спираль
                                    stLine = AlignSCS.StartStation
                                    enLine = AlignSCS.EndStation
                                    '=================================================================================================
                                ElseIf TypeOf myAe Is AlignmentSpiral Then
                                    Dim AlignS3 As AlignmentSpiral = myAe 'отдельная спираль
                                    stLine = AlignS3.StartStation
                                    enLine = AlignS3.EndStation
                                End If
'=================================================================================================
'спираль дуга спираль
                            Case AlignmentEntityType.SpiralCurveSpiral
                                Dim AlignSCS As AlignmentSCS = myAe
                                Dim Spin As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Spin = AlignSCS.SpiralIn
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Spin) = False Then
                                    If Spin.StartStation <= PK And Spin.EndStation >= PK Then
                                        stLine = Spin.StartStation
                                        enLine = Spin.EndStation
                                    End If
                                End If
                                Dim Arc1 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc1 = AlignSCS.Arc
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc1) = False Then
                                    If Arc1.StartStation <= PK And Arc1.EndStation >= PK Then
                                        stLine = Arc1.StartStation
                                        enLine = Arc1.EndStation
                                    End If
                                End If
                                Dim SpOut As AlignmentSubEntitySpiral = Nothing
                                Try
                                    SpOut = AlignSCS.SpiralOut
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(SpOut) = False Then
                                    If SpOut.StartStation <= PK And SpOut.EndStation >= PK Then
                                        stLine = SpOut.StartStation
                                        enLine = SpOut.EndStation
                                    End If
                                End If
'=============================================================================================================================
'переходная кривая прямая переходная кривая
                            Case AlignmentEntityType.SpiralLineSpiral
                                Dim AlignSLS As AlignmentSTS = myAe
                                Dim Spin As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Spin = AlignSLS.SpiralIn
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Spin) = False Then
                                    If Spin.StartStation <= PK And Spin.EndStation >= PK Then
                                        stLine = Spin.StartStation
                                        enLine = Spin.EndStation
                                    End If
                                End If
                                Dim Line1 As AlignmentSubEntityLine = Nothing
                                Try
                                    Line1 = AlignSLS.Tangent
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Line1) = False Then
                                    If Line1.StartStation <= PK And Line1.EndStation >= PK Then
                                        stLine = Line1.StartStation
                                        enLine = Line1.EndStation
                                    End If
                                End If
                                Dim SpOut As AlignmentSubEntitySpiral = Nothing
                                Try
                                    SpOut = AlignSLS.SpiralOut
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(SpOut) = False Then
                                    If SpOut.StartStation <= PK And SpOut.EndStation >= PK Then
                                        stLine = SpOut.StartStation
                                        enLine = SpOut.EndStation
                                    End If
                                End If
'=========================================================================================================================
'переходная кривая прямая
                            Case AlignmentEntityType.SpiralLine
                                Dim AlignSL As AlignmentSTS = myAe
                                Dim Spin As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Spin = AlignSL.SpiralIn
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Spin) = False Then
                                    If Spin.StartStation <= PK And Spin.EndStation >= PK Then
                                        stLine = Spin.StartStation
                                        enLine = Spin.EndStation
                                    End If
                                End If
                                Dim Line1 As AlignmentSubEntityLine = Nothing
                                Try
                                    Line1 = AlignSL.Tangent
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Line1) = False Then
                                    If Line1.StartStation <= PK And Line1.EndStation >= PK Then
                                        stLine = Line1.StartStation
                                        enLine = Line1.EndStation
                                    End If
                                End If
'======================================================================================================================
'прямая спираль 
                            Case AlignmentEntityType.LineSpiral
                                Dim AlignLS As AlignmentSTS = myAe
                                Dim Line1 As AlignmentSubEntityLine = Nothing
                                Try
                                    Line1 = AlignLS.Tangent
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Line1) = False Then
                                    If Line1.StartStation <= PK And Line1.EndStation >= PK Then
                                        stLine = Line1.StartStation
                                        enLine = Line1.EndStation
                                    End If
                                End If
                                Dim SpOut As AlignmentSubEntitySpiral = Nothing
                                Try
                                    SpOut = AlignLS.SpiralOut
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(SpOut) = False Then
                                    If SpOut.StartStation <= PK And SpOut.EndStation >= PK Then
                                        stLine = SpOut.StartStation
                                        enLine = SpOut.EndStation
                                    End If
                                End If
'===================================================================================================================================
'переходная кривая дуга 
                            Case AlignmentEntityType.SpiralCurve
                                Dim AlignSC As AlignmentSCS = myAe
                                Dim Spin As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Spin = AlignSC.SpiralIn
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Spin) = False Then
                                    If Spin.StartStation <= PK And Spin.EndStation >= PK Then
                                        stLine = Spin.StartStation
                                        enLine = Spin.EndStation
                                    End If
                                End If
                                Dim Arc1 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc1 = AlignSC.Arc
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc1) = False Then
                                    If Arc1.StartStation <= PK And Arc1.EndStation >= PK Then
                                        stLine = Arc1.StartStation
                                        enLine = Arc1.EndStation
                                    End If
                                End If
'==========================================================================================================================================
                            Case AlignmentEntityType.CurveSpiral 'кривая переходная кривая
                                Dim AlignCS As AlignmentSCS = myAe
                                Dim Arc1 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc1 = AlignCS.Arc
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc1) = False Then
                                    If Arc1.StartStation <= PK And Arc1.EndStation >= PK Then
                                        stLine = Arc1.StartStation
                                        enLine = Arc1.EndStation
                                    End If
                                End If
                                Dim SpOut As AlignmentSubEntitySpiral = AlignCS.SpiralOut
                                Try
                                    SpOut = AlignCS.SpiralOut
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(SpOut) = False Then
                                    If SpOut.StartStation <= PK And SpOut.EndStation >= PK Then
                                        stLine = SpOut.StartStation
                                        enLine = SpOut.EndStation
                                    End If
                                End If
'==============================================================================================================================================
'переходная кривая переходная кривая дуга переходная кривая переходная кривая
                            Case AlignmentEntityType.SpiralSpiralCurveSpiralSpiral
                                Dim AlignSSCSS As AlignmentSSCSS = myAe
                                Dim Sp1 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp1 = AlignSSCSS.Spiral1
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp1) = False Then
                                    If Sp1.StartStation <= PK And Sp1.EndStation >= PK Then
                                        stLine = Sp1.StartStation
                                        enLine = Sp1.EndStation
                                    End If
                                End If
                                Dim Sp2 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp2 = AlignSSCSS.Spiral2
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp2) = False Then
                                    If Sp2.StartStation <= PK And Sp2.EndStation >= PK Then
                                        stLine = Sp2.StartStation
                                        enLine = Sp2.EndStation
                                    End If
                                End If
                                Dim Arc1 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc1 = AlignSSCSS.Arc
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc1) = False Then
                                    If Arc1.StartStation <= PK And Arc1.EndStation >= PK Then
                                        stLine = Arc1.StartStation
                                        enLine = Arc1.EndStation
                                    End If
                                End If
                                Dim Sp3 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp3 = AlignSSCSS.Spiral3
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp3) = False Then
                                    If Sp3.StartStation <= PK And Sp3.EndStation >= PK Then
                                        stLine = Sp3.StartStation
                                        enLine = Sp3.EndStation
                                    End If
                                End If
                                Dim Sp4 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp4 = AlignSSCSS.Spiral4
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp4) = False Then
                                    If Sp4.StartStation <= PK And Sp4.EndStation >= PK Then
                                        stLine = Sp4.StartStation
                                        enLine = Sp4.EndStation
                                    End If
                                End If
'====================================================================================================================================
'дуга дуга обратная дуга
                            Case AlignmentEntityType.CurveCurveReverseCurve
                                Dim AlignCCRC As AlignmentCCRC = myAe
                                Dim Arc1 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc1 = AlignCCRC.Arc1
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc1) = False Then
                                    If Arc1.StartStation <= PK And Arc1.EndStation >= PK Then
                                        stLine = Arc1.StartStation
                                        enLine = Arc1.EndStation
                                    End If
                                End If
                                Dim Arc2 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc2 = AlignCCRC.Arc2
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc2) = False Then
                                    If Arc2.StartStation <= PK And Arc2.EndStation >= PK Then
                                        stLine = Arc2.StartStation
                                        enLine = Arc2.EndStation
                                    End If
                                End If
                                Dim Arc3 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc3 = AlignCCRC.Arc3
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc3) = False Then
                                    If Arc3.StartStation <= PK And Arc3.EndStation >= PK Then
                                        stLine = Arc3.StartStation
                                        enLine = Arc3.EndStation
                                    End If
                                End If
'==============================================================================================================================
'дуга обратная дуга
                            Case AlignmentEntityType.CurveReverseCurve
                                Dim AlignCRC As AlignmentCRC = myAe
                                Dim Arc1 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc1 = AlignCRC.Arc1
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc1) = False Then
                                    If Arc1.StartStation <= PK And Arc1.EndStation >= PK Then
                                        stLine = Arc1.StartStation
                                        enLine = Arc1.EndStation
                                    End If
                                End If
                                Dim Arc2 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc2 = AlignCRC.Arc2
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc2) = False Then
                                    If Arc2.StartStation <= PK And Arc2.EndStation >= PK Then
                                        stLine = Arc2.StartStation
                                        enLine = Arc2.EndStation
                                    End If
                                End If
'==================================================================================================================================
'кривавая спираль спираль
                            Case AlignmentEntityType.CurveSpiralSpiral
                                Dim AlignCSS As AlignmentSSCSS = myAe
                                Dim Arc1 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc1 = AlignCSS.Arc
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc1) = False Then
                                    If Arc1.StartStation <= PK And Arc1.EndStation >= PK Then
                                        stLine = Arc1.StartStation
                                        enLine = Arc1.EndStation
                                    End If
                                End If
                                Dim Sp3 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp3 = AlignCSS.Spiral3
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp3) = False Then
                                    If Sp3.StartStation <= PK And Sp3.EndStation >= PK Then
                                        stLine = Sp3.StartStation
                                        enLine = Sp3.EndStation
                                    End If
                                End If
                                Dim Sp4 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp4 = AlignCSS.Spiral4
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp4) = False Then
                                    If Sp4.StartStation <= PK And Sp4.EndStation >= PK Then
                                        stLine = Sp4.StartStation
                                        enLine = Sp4.EndStation
                                    End If
                                End If
'==========================================================================================================================
'спираль спираль
                            Case AlignmentEntityType.SpiralSpiral
                                Dim Sp1 As AlignmentSubEntitySpiral = Nothing
                                Dim Sp2 As AlignmentSubEntitySpiral = Nothing
                                If TypeOf myAe Is AlignmentSTS Then
                                    Dim AlignCS1 As AlignmentSTS = myAe
                                    Try
                                        Sp1 = AlignCS1.SpiralIn
                                    Catch ex As System.InvalidOperationException
                                    End Try
                                    Try
                                        Sp2 = AlignCS1.SpiralOut
                                    Catch ex As System.InvalidOperationException
                                    End Try
                                ElseIf TypeOf myAe Is AlignmentSCS Then
                                    Dim AlignCS2 As AlignmentSCS = myAe
                                    Try
                                        Sp1 = AlignCS2.SpiralIn
                                    Catch ex As System.InvalidOperationException
                                    End Try
                                    Try
                                        Sp2 = AlignCS2.SpiralOut
                                    Catch ex As System.InvalidOperationException
                                    End Try
                                End If
                                If IsNothing(Sp1) = False Then
                                    If Sp1.StartStation <= PK And Sp1.EndStation >= PK Then
                                        stLine = Sp1.StartStation
                                        enLine = Sp1.EndStation
                                    End If
                                End If
                                If IsNothing(Sp2) = False Then
                                    If Sp2.StartStation <= PK And Sp2.EndStation >= PK Then
                                        stLine = Sp2.StartStation
                                        enLine = Sp2.EndStation
                                    End If
                                End If
'=============================================================================================================================
'спираль спираль дуга
                            Case AlignmentEntityType.SpiralSpiralCurve
                                Dim AlignSSC As AlignmentSSCSS = myAe
                                Dim Sp1 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp1 = AlignSSC.Spiral1
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp1) = False Then
                                    If Sp1.StartStation <= PK And Sp1.EndStation >= PK Then
                                        stLine = Sp1.StartStation
                                        enLine = Sp1.EndStation
                                    End If
                                End If
                                Dim Sp2 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp2 = AlignSSC.Spiral2
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp2) = False Then
                                    If Sp2.StartStation <= PK And Sp2.EndStation >= PK Then
                                        stLine = Sp2.StartStation
                                        enLine = Sp2.EndStation
                                    End If
                                End If
                                Dim Arc1 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc1 = AlignSSC.Arc
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc1) = False Then
                                    If Arc1.StartStation <= PK And Arc1.EndStation >= PK Then
                                        stLine = Arc1.StartStation
                                        enLine = Arc1.EndStation
                                    End If
                                End If
'================================================================================================================================
'спираль дуга спираль спираль дуга спираль
                            Case AlignmentEntityType.SpiralCurveSpiralSpiralCurveSpiral
                                Dim AlignSCSSCS As AlignmentSCSSCS = myAe
                                Dim Sp1 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp1 = AlignSCSSCS.Spiral1
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp1) = False Then
                                    If Sp1.StartStation <= PK And Sp1.EndStation >= PK Then
                                        stLine = Sp1.StartStation
                                        enLine = Sp1.EndStation
                                    End If
                                End If
                                Dim Arc1 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc1 = AlignSCSSCS.Arc1
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc1) = False Then
                                    If Arc1.StartStation <= PK And Arc1.EndStation >= PK Then
                                        stLine = Arc1.StartStation
                                        enLine = Arc1.EndStation
                                    End If
                                End If
                                Dim Sp2 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp2 = AlignSCSSCS.Spiral2
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp2) = False Then
                                    If Sp2.StartStation <= PK And Sp2.EndStation >= PK Then
                                        stLine = Sp2.StartStation
                                        enLine = Sp2.EndStation
                                    End If
                                End If
                                Dim Sp3 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp3 = AlignSCSSCS.Spiral3
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp3) = False Then
                                    If Sp3.StartStation <= PK And Sp3.EndStation >= PK Then
                                        stLine = Sp3.StartStation
                                        enLine = Sp3.EndStation
                                    End If
                                End If
                                Dim Arc2 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc2 = AlignSCSSCS.Arc2
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc2) = False Then
                                    If Arc2.StartStation <= PK And Arc2.EndStation >= PK Then
                                        stLine = Arc2.StartStation
                                        enLine = Arc2.EndStation
                                    End If
                                End If
                                Dim Sp4 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp4 = AlignSCSSCS.Spiral4
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp4) = False Then
                                    If Sp4.StartStation <= PK And Sp4.EndStation >= PK Then
                                        stLine = Sp4.StartStation
                                        enLine = Sp4.EndStation
                                    End If
                                End If
'==================================================================================================================================================
'спираль дуга спираль дуга спираль
                            Case AlignmentEntityType.SpiralCurveSpiralCurveSpiral
                                Dim AlignSCSCS As AlignmentSCSCS = myAe
                                Dim Sp1 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp1 = AlignSCSCS.Spiral1
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp1) = False Then
                                    If Sp1.StartStation <= PK And Sp1.EndStation >= PK Then
                                        stLine = Sp1.StartStation
                                        enLine = Sp1.EndStation
                                    End If
                                End If
                                Dim Arc1 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc1 = AlignSCSCS.Arc1
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc1) = False Then
                                    If Arc1.StartStation <= PK And Arc1.EndStation >= PK Then
                                        stLine = Arc1.StartStation
                                        enLine = Arc1.EndStation
                                    End If
                                End If
                                Dim Sp2 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp2 = AlignSCSCS.Spiral2
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp2) = False Then
                                    If Sp2.StartStation <= PK And Sp2.EndStation >= PK Then
                                        stLine = Sp2.StartStation
                                        enLine = Sp2.EndStation
                                    End If
                                End If
                                Dim Arc2 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc2 = AlignSCSCS.Arc2
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc2) = False Then
                                    If Arc2.StartStation <= PK And Arc2.EndStation >= PK Then
                                        stLine = Arc2.StartStation
                                        enLine = Arc2.EndStation
                                    End If
                                End If
                                Dim Sp3 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    Sp3 = AlignSCSCS.Spiral3
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Sp3) = False Then
                                    If Sp3.StartStation <= PK And Sp3.EndStation >= PK Then
                                        stLine = Sp3.StartStation
                                        enLine = Sp3.EndStation
                                    End If
                                End If
 '=================================================================================================================                            
                            Case AlignmentEntityType.MultipleSegments
                                Dim AlignCS As AlignmentMultipleSegments = myAe
                            'структура не понятна
'=================================================================================================================
                            Case AlignmentEntityType.CurveLineCurve
                                Dim AlignCLC As AlignmentCTC = myAe
                                Dim Arc1 As AlignmentSubEntityArc = AlignCLC.Arc1
                                Try
                                    Arc1 = AlignCLC.Arc1
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc1) = False Then
                                    If Arc1.StartStation <= PK And Arc1.EndStation >= PK Then
                                        stLine = Arc1.StartStation
                                        enLine = Arc1.EndStation
                                    End If
                                End If
                                Dim Line1 As AlignmentSubEntityLine = Nothing
                                Try
                                    Line1 = AlignCLC.Tangent
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Line1) = False Then
                                    If Line1.StartStation <= PK And Line1.EndStation >= PK Then
                                        stLine = Line1.StartStation
                                        enLine = Line1.EndStation
                                    End If
                                End If
                                Dim Arc2 As AlignmentSubEntityArc = Nothing
                                Try
                                    Arc2 = AlignCLC.Arc1
                                Catch ex As System.InvalidOperationException
                                End Try
                                If IsNothing(Arc2) = False Then
                                    If Arc2.StartStation <= PK And Arc2.EndStation >= PK Then
                                        stLine = Arc2.StartStation
                                        enLine = Arc2.EndStation
                                    End If
                                End If
                        End Select
                        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'вставка блока
                        If stLine = 0 And enLine = 0 Then Continue For
                        If stLine > enLine Then
                            Dim tempSt As Double = stLine
                            stLine = enLine
                            enLine = tempSt
                        End If
                        If Math.Abs(stLine - enLine) <= interval Then Continue For
                        If PK >= stLine And PK <= enLine Then
                            stLine = stLine + interval
                            For i As Double = stLine To enLine Step interval
                                If i = enLine Then Exit For
                                Dim X As Double = -1
                                Dim Y As Double = -1
                                Try
                                    Align.PointLocation(i, UserOffset, X, Y)
                                    Dim pt3d As Point3d = New Point3d(X, Y, 0)
                                    Dim blk As BlockReference = AcadFunction.FuncInsBlkByArrayProperties(pt3d, ArrayPropBlk, rot)
                                Catch ex As Autodesk.Civil.PointNotOnEntityException
                                End Try
                            Next i
                        End If
                    Next
                Loop
                tr.Commit()
            End Using
        Catch
        End Try
    End Sub
    '=============================================================================================
    'поставить пикеты в мультивыноске
    <CommandMethod("AddMleaderPK")>
    Public Sub AddMleaderPK()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ссылка на чертеж .dwg
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
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
        Dim PutchTemplateDir As String = ""
        Dim PutchTemplateDwg As String = ""
        '========================================================================================
        'загружаем пути к сборкам
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        '=============================================================================================================
        'проверяем и загружаем все стили из шаблона
        '=============================================================================================================
        Dim ArrayFilesDirTemlates As String() = Nothing
        Dim CountPutchTemplateDir As Integer = 0
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            PutchTemplateDir = parentPatch2 & "\FileResources\Sample\"
            For Each fileDwg As String In Directory.GetFiles(PutchTemplateDir)
                Try
                    Dim File1 As String = Path.GetFileName(fileDwg)
                    Dim ext As String = Path.GetExtension(fileDwg)
                    If ext Like ".dwg" Then
                        PutchTemplateDwg = fileDwg
                        Exit For
                    End If
                Catch ex As ArgumentException
                End Try
            Next
        End If
        'копируем стиль мультивыноски
        Dim DeskMleader As String = "Пикет"
        Dim nameStyleMleader As String = "Standard"
        Dim idStyleMleader As ObjectId = db.MLeaderstyle
        Dim nameLayerMleader As String = "0"
        Dim indexColorMLeader As Integer = 256
        Dim TypeLineNLeader As String = "Continuous"
        Dim WidthLineMLeader As Double = 20
        Dim ArrowSize As Double = 2
        Dim StyleTextMleader As String = "Standard"
        Dim idTextStyle As ObjectId = Nothing
        Dim HText As Double = 2.5
        Dim PSTable As String = ""
        Dim idTablePs As ObjectId = Nothing
        Dim BaskGround As Boolean = True
        Dim ArrayPropertiesMLeader As String(,) = Nothing
        Dim tr0 As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr0
            If File.Exists(PutchTemplateDwg) = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskMleader, ArrayPropertiesMLeader, "Мультивыноска", "Трасса автомобильной дороги")
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'делаем разбор массива
                If IsArray(ArrayPropertiesMLeader) = True Then
                    For i As Integer = 0 To ArrayPropertiesMLeader.GetUpperBound(1)
                        Dim NameInsObject As String = ArrayPropertiesMLeader(1, i)
                        If ArrayPropertiesMLeader(0, i) Like "Name" Then
                            If IsNothing(NameInsObject) = False Then
                                Dim mlStyles As DBDictionary = tr0.GetObject(db.MLeaderStyleDictionaryId, OpenMode.ForRead)
                                If mlStyles.Contains(NameInsObject) = False Then
                                    Dim boolCopyBlk As Boolean = AcadFunction.FuncCreateMLeaderStyleDestDraw(PutchTemplateDwg, NameInsObject)
                                    If boolCopyBlk = True Then
                                        idStyleMleader = mlStyles.GetAt(NameInsObject)
                                    End If
                                Else
                                    idStyleMleader = mlStyles.GetAt(NameInsObject)
                                End If
                            End If
                        ElseIf ArrayPropertiesMLeader(0, i) Like "Layer" Then
                            Dim LayerinsBlk As String = ArrayPropertiesMLeader(1, i)
                            If IsNothing(LayerinsBlk) = False Then
                                Dim acLyrTbl As LayerTable = tr0.GetObject(db.LayerTableId, OpenMode.ForRead)
                                If acLyrTbl.Has(LayerinsBlk) = False Then
                                    Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PutchTemplateDwg, LayerinsBlk)
                                    If boolCopyLayer = True Then
                                        nameLayerMleader = LayerinsBlk
                                    End If
                                Else
                                    nameLayerMleader = LayerinsBlk
                                End If
                            End If
                        ElseIf ArrayPropertiesMLeader(0, i) Like "PS" Then
                            PSTable = ArrayPropertiesMLeader(1, i)
                            If IsNothing(PSTable) = False Then
                                Dim dictPropSetDefTemp = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(db)
                                If dictPropSetDefTemp.Has(PSTable, tr0) = False Then
                                    Dim boolCopyPSTableTemp As Boolean = AcadFunction.FuncCopyTablePSDestDraw(PutchTemplateDwg, PSTable)
                                End If
                            End If
                            idTablePs = FuncPropertySet.FuncReturnIDDictSetProperties(PSTable)
                        ElseIf ArrayPropertiesMLeader(0, i) Like "Color" Then
                            Dim IndexColor As String = ArrayPropertiesMLeader(1, i)
                            If IndexColor Like "ПоСлою" Then
                                indexColorMLeader = 256
                            ElseIf IndexColor Like "ПоБлоку" Then
                                indexColorMLeader = 0
                            ElseIf IndexColor Like "ByLayer" Then
                                indexColorMLeader = 256
                            ElseIf IndexColor Like "ByBlock" Then
                                indexColorMLeader = 0
                            ElseIf Val(IndexColor) < 1 Then
                                indexColorMLeader = 0
                            ElseIf Val(IndexColor) > 255 Then
                                indexColorMLeader = 256
                            Else
                                indexColorMLeader = Val(IndexColor)
                            End If
                        ElseIf ArrayPropertiesMLeader(0, i) Like "LineType" Then
                            Dim strTypeLine As String = ArrayPropertiesMLeader(1, i)
                            If strTypeLine Like "ПоСлою" Then
                                TypeLineNLeader = "ПоСлою"
                            ElseIf strTypeLine Like "ПоБлоку" Then
                                TypeLineNLeader = "ПоБлоку"
                            ElseIf strTypeLine Like "ByLayer" Then
                                TypeLineNLeader = "ПоСлою"
                            ElseIf strTypeLine Like "ByBlock" Then
                                TypeLineNLeader = "ПоБлоку"
                            Else
                                Dim acLineTypTbl As LinetypeTable = tr0.GetObject(db.LinetypeTableId, OpenMode.ForRead)
                                If acLineTypTbl.Has(strTypeLine) = False Then
                                    'копируем тип линии из файла-шаблона
                                    Dim boolTypeline As Boolean = AcadFunction.FuncCreateTypeLineDestDraw(PutchTemplateDwg, strTypeLine)
                                End If
                                If acLineTypTbl.Has(strTypeLine) = True Then
                                    TypeLineNLeader = strTypeLine
                                End If
                            End If
                        ElseIf ArrayPropertiesMLeader(0, i) Like "LineWidth" Then
                            WidthLineMLeader = Val(ArrayPropertiesMLeader(1, i))
                        ElseIf ArrayPropertiesMLeader(0, i) Like "ArrowSize" Then
                            ArrowSize = Val(ArrayPropertiesMLeader(1, i))
                        ElseIf ArrayPropertiesMLeader(0, i) Like "TextStyle" Then
                            Dim strTextStyle As String = ArrayPropertiesMLeader(1, i)
                            'создаем текстовый стиль для точек
                            Dim st As TextStyleTable = CType(tr0.GetObject(db.TextStyleTableId, OpenMode.ForWrite, False), TextStyleTable)
                            If st.Has(strTextStyle) = False Then
                                Dim crTextstyle As Boolean = AcadFunction.FuncCreateTextStyleDestDraw(PutchTemplateDwg, strTextStyle, False)
                            End If
                            If st.Has(strTextStyle) = True Then
                                StyleTextMleader = strTextStyle
                                idTextStyle = AcadFunction.FuncReturnIDTextStylebyName(StyleTextMleader)
                            End If
                        ElseIf ArrayPropertiesMLeader(0, i) Like "TextHeight" Then
                            HText = Val(ArrayPropertiesMLeader(1, i))
                        End If
                    Next i
                End If
            End If
            tr0.Commit()
        End Using
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '=============================================================================================================
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr
            Dim ppr1 As PromptEntityResult = ed.GetEntity("Выберите трассу: ")
            Dim ent1 As Autodesk.Civil.DatabaseServices.Entity = tr.GetObject(ppr1.ObjectId, OpenMode.ForRead)
            If TypeOf ent1 Is Alignment Then
                Dim align As Alignment = ent1
                Do
                    Dim ptSelect As PromptPointResult = ed.GetPoint("Укажите точку: ")
                    If ptSelect.Status = PromptStatus.OK Then
                        Dim x As Double = ptSelect.Value.X
                        Dim y As Double = ptSelect.Value.Y
                        Dim insMleader As Point2d = New Point2d(x, y)
                        Dim st As Double = -1
                        Dim off As Double = -1
                        align.StationOffset(x, y, st, off)
                        If st <> -1 Then
                            Dim PKStr As String = align.GetStationStringWithEquations(st)
                            Dim idlinsMline As ObjectId = AcadFunction.FuncDrawMleader(insMleader, "ПК" & PKStr, 5, 5, nameLayerMleader, HText, idStyleMleader, idTextStyle,, True)
                            If IsNothing(idlinsMline) = False Then
                                If idlinsMline.IsNull = False Then
                                    Dim acMleader As MLeader = idlinsMline.GetObject(OpenMode.ForWrite)
                                    acMleader.ArrowSize = ArrowSize
                                    acMleader.ColorIndex = indexColorMLeader
                                    acMleader.LineWeight = WidthLineMLeader
                                    '===================================================================================================
                                    'присваиваем дополнительные свойства
                                    '=================================================================================
                                    'формируем массив для записи
                                    Dim ArrayOBJ1 As String(,) = Nothing
                                    ReDim Preserve ArrayOBJ1(2, 2)
                                    ArrayOBJ1(0, 0) = "OBJECT"
                                    ArrayOBJ1(1, 0) = "LABEL"
                                    ArrayOBJ1(2, 0) = DeskMleader

                                    ArrayOBJ1(0, 1) = "OBJECT"
                                    ArrayOBJ1(1, 1) = "TEXT1"
                                    ArrayOBJ1(2, 1) = "ПК" & PKStr

                                    ArrayOBJ1(0, 2) = "OBJECT"
                                    ArrayOBJ1(1, 2) = "TEXT2"
                                    ArrayOBJ1(2, 2) = align.Name
                                    '=======================================================================================================
                                    If IsNothing(idTablePs) = False Then
                                        If idTablePs.IsNull = False Then
                                            Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(acMleader, idTablePs)
                                            Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acMleader, ArrayOBJ1)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        Exit Do
                    End If
                Loop
                tr.Commit()
            End If
        End Using
    End Sub
    '=============================================================================================
    'поставить пикеты в мультивыноске
    <CommandMethod("AddMleaderSQRT")>
    Public Sub AddMleaderSQRT()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ссылка на чертеж .dwg
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
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
        Dim PutchTemplateDir As String = ""
        Dim PutchTemplateDwg As String = ""
        '========================================================================================
        'загружаем пути к сборкам
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        'отделяем имя файла от папки
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
        'определяем имя родительской папки для папки где находится сборка
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
        '=============================================================================================================
        'проверяем и загружаем все стили из шаблона
        '=============================================================================================================
        Dim ArrayFilesDirTemlates As String() = Nothing
        Dim CountPutchTemplateDir As Integer = 0
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            PutchTemplateDir = parentPatch2 & "\FileResources\Sample\"
            For Each fileDwg As String In Directory.GetFiles(PutchTemplateDir)
                Try
                    Dim File1 As String = Path.GetFileName(fileDwg)
                    Dim ext As String = Path.GetExtension(fileDwg)
                    If ext Like ".dwg" Then
                        PutchTemplateDwg = fileDwg
                        Exit For
                    End If
                Catch ex As ArgumentException
                End Try
            Next
        End If
        'копируем стиль мультивыноски
        Dim DeskMleader As String = "Площадь объекта"
        Dim nameStyleMleader As String = "Standard"
        Dim idStyleMleader As ObjectId = db.MLeaderstyle
        Dim nameLayerMleader As String = "0"
        Dim indexColorMLeader As Integer = 256
        Dim TypeLineNLeader As String = "Continuous"
        Dim WidthLineMLeader As Double = 20
        Dim ArrowSize As Double = 2
        Dim StyleTextMleader As String = "Standard"
        Dim idTextStyle As ObjectId = Nothing
        Dim HText As Double = 2.5
        Dim PSTable As String = ""
        Dim idTablePs As ObjectId = Nothing
        Dim BaskGround As Boolean = True
        Dim ArrayPropertiesMLeader As String(,) = Nothing
        Dim tr0 As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        Using tr0
            If File.Exists(PutchTemplateDwg) = True Then
                Dim boolFindObj As Boolean = AcadFunction.FuncReadPropertyObjectDestDWG(PutchTemplateDwg, DeskMleader, ArrayPropertiesMLeader, "Мультивыноска", "Трасса автомобильной дороги")
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'делаем разбор массива
                If IsArray(ArrayPropertiesMLeader) = True Then
                    For i As Integer = 0 To ArrayPropertiesMLeader.GetUpperBound(1)
                        Dim NameInsObject As String = ArrayPropertiesMLeader(1, i)
                        If ArrayPropertiesMLeader(0, i) Like "Name" Then
                            If IsNothing(NameInsObject) = False Then
                                Dim mlStyles As DBDictionary = tr0.GetObject(db.MLeaderStyleDictionaryId, OpenMode.ForRead)
                                If mlStyles.Contains(NameInsObject) = False Then
                                    Dim boolCopyBlk As Boolean = AcadFunction.FuncCreateMLeaderStyleDestDraw(PutchTemplateDwg, NameInsObject)
                                    If boolCopyBlk = True Then
                                        idStyleMleader = mlStyles.GetAt(NameInsObject)
                                    End If
                                Else
                                    idStyleMleader = mlStyles.GetAt(NameInsObject)
                                End If
                            End If
                        ElseIf ArrayPropertiesMLeader(0, i) Like "Layer" Then
                            Dim LayerinsBlk As String = ArrayPropertiesMLeader(1, i)
                            If IsNothing(LayerinsBlk) = False Then
                                Dim acLyrTbl As LayerTable = tr0.GetObject(db.LayerTableId, OpenMode.ForRead)
                                If acLyrTbl.Has(LayerinsBlk) = False Then
                                    Dim boolCopyLayer As Boolean = AcadFunction.FuncCreateLayerDestDraw(PutchTemplateDwg, LayerinsBlk)
                                    If boolCopyLayer = True Then
                                        nameLayerMleader = LayerinsBlk
                                    End If
                                Else
                                    nameLayerMleader = LayerinsBlk
                                End If
                            End If
                        ElseIf ArrayPropertiesMLeader(0, i) Like "PS" Then
                            PSTable = ArrayPropertiesMLeader(1, i)
                            If IsNothing(PSTable) = False Then
                                Dim dictPropSetDefTemp = New Autodesk.AEC.PropertyData.DatabaseServices.DictionaryPropertySetDefinitions(db)
                                If dictPropSetDefTemp.Has(PSTable, tr0) = False Then
                                    Dim boolCopyPSTableTemp As Boolean = AcadFunction.FuncCopyTablePSDestDraw(PutchTemplateDwg, PSTable)
                                End If
                            End If
                            idTablePs = FuncPropertySet.FuncReturnIDDictSetProperties(PSTable)
                        ElseIf ArrayPropertiesMLeader(0, i) Like "Color" Then
                            Dim IndexColor As String = ArrayPropertiesMLeader(1, i)
                            If IndexColor Like "ПоСлою" Then
                                indexColorMLeader = 256
                            ElseIf IndexColor Like "ПоБлоку" Then
                                indexColorMLeader = 0
                            ElseIf IndexColor Like "ByLayer" Then
                                indexColorMLeader = 256
                            ElseIf IndexColor Like "ByBlock" Then
                                indexColorMLeader = 0
                            ElseIf Val(IndexColor) < 1 Then
                                indexColorMLeader = 0
                            ElseIf Val(IndexColor) > 255 Then
                                indexColorMLeader = 256
                            Else
                                indexColorMLeader = Val(IndexColor)
                            End If
                        ElseIf ArrayPropertiesMLeader(0, i) Like "LineType" Then
                            Dim strTypeLine As String = ArrayPropertiesMLeader(1, i)
                            If strTypeLine Like "ПоСлою" Then
                                TypeLineNLeader = "ПоСлою"
                            ElseIf strTypeLine Like "ПоБлоку" Then
                                TypeLineNLeader = "ПоБлоку"
                            ElseIf strTypeLine Like "ByLayer" Then
                                TypeLineNLeader = "ПоСлою"
                            ElseIf strTypeLine Like "ByBlock" Then
                                TypeLineNLeader = "ПоБлоку"
                            Else
                                Dim acLineTypTbl As LinetypeTable = tr0.GetObject(db.LinetypeTableId, OpenMode.ForRead)
                                If acLineTypTbl.Has(strTypeLine) = False Then
                                    'копируем тип линии из файла-шаблона
                                    Dim boolTypeline As Boolean = AcadFunction.FuncCreateTypeLineDestDraw(PutchTemplateDwg, strTypeLine)
                                End If
                                If acLineTypTbl.Has(strTypeLine) = True Then
                                    TypeLineNLeader = strTypeLine
                                End If
                            End If
                        ElseIf ArrayPropertiesMLeader(0, i) Like "LineWidth" Then
                            WidthLineMLeader = Val(ArrayPropertiesMLeader(1, i))
                        ElseIf ArrayPropertiesMLeader(0, i) Like "ArrowSize" Then
                            ArrowSize = Val(ArrayPropertiesMLeader(1, i))
                        ElseIf ArrayPropertiesMLeader(0, i) Like "TextStyle" Then
                            Dim strTextStyle As String = ArrayPropertiesMLeader(1, i)
                            'создаем текстовый стиль для точек
                            Dim st As TextStyleTable = CType(tr0.GetObject(db.TextStyleTableId, OpenMode.ForWrite, False), TextStyleTable)
                            If st.Has(strTextStyle) = False Then
                                Dim crTextstyle As Boolean = AcadFunction.FuncCreateTextStyleDestDraw(PutchTemplateDwg, strTextStyle, False)
                            End If
                            If st.Has(strTextStyle) = True Then
                                StyleTextMleader = strTextStyle
                                idTextStyle = AcadFunction.FuncReturnIDTextStylebyName(StyleTextMleader)
                            End If
                        ElseIf ArrayPropertiesMLeader(0, i) Like "TextHeight" Then
                            HText = Val(ArrayPropertiesMLeader(1, i))
                        End If
                    Next i
                End If
            End If
            tr0.Commit()
        End Using
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '=============================================================================================================
        Do
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Dim ppr1 As PromptEntityResult = ed.GetEntity("Выберите объект: ")
                If ppr1.Status = PromptStatus.OK Then
                    Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ppr1.ObjectId, OpenMode.ForRead)
                    Dim insPoint As Point3d = ppr1.PickedPoint
                    Dim idlinsMline As ObjectId = AcadFunction.FuncDrawMleader(insPoint.Convert2d(New Plane), "", 5, 5, nameLayerMleader, HText, idStyleMleader, idTextStyle,, True)
                    If IsNothing(idlinsMline) = False Then
                        If idlinsMline.IsNull = False Then
                            Dim acMleader As MLeader = idlinsMline.GetObject(OpenMode.ForWrite)
                            acMleader.ArrowSize = ArrowSize
                            acMleader.ColorIndex = indexColorMLeader
                            acMleader.LineWeight = WidthLineMLeader
                            Dim acMtext As String = acMleader.MText.Contents
                            '============================================================================================
                            'разбираем строку
                            Dim id As ObjectId = acMleader.GetField() 'смотрим старый атрибут
                            Dim fd As Field = tr.GetObject(id, OpenMode.ForWrite) 'читаем его
                            Dim CodStrFiled As String = fd.GetFieldCode 'получаем код поля
                            Dim pos2 As Integer = CodStrFiled.IndexOf("Area")
                            If pos2 > -1 Then
                                Dim StrId As Long = ent1.Id.OldIdPtr
                                CodStrFiled = "\A1;S=" & "%<\AcObjProp Object(%<\_ObjId " & StrId & ">%).Area \f " & """" & "%lu2%pr2" & """" & ">%" & "м{\H0.7x;\S2^;}"
                                fd.SetFieldCode(CodStrFiled)
                                fd.Evaluate()
                            End If
                            '===================================================================================================
                            'присваиваем дополнительные свойства
                            '=================================================================================
                            'формируем массив для записи
                            Dim ArrayOBJ1 As String(,) = Nothing
                            ReDim Preserve ArrayOBJ1(2, 2)
                            ArrayOBJ1(0, 0) = "OBJECT"
                            ArrayOBJ1(1, 0) = "LABEL"
                            ArrayOBJ1(2, 0) = DeskMleader
                            '=======================================================================================================
                            If IsNothing(idTablePs) = False Then
                                If idTablePs.IsNull = False Then
                                    Dim boolCreateTbl As Boolean = FuncPropertySet.FuncAddDataSetProperties(acMleader, idTablePs)
                                    Dim boolWriteDAta As Boolean = FuncPropertySet.FuncWriteAllDataSetProperties(acMleader, ArrayOBJ1)
                                End If
                            End If
                        End If
                    End If
                Else
                    Exit Do
                End If
                tr.Commit()
            End Using
        Loop
    End Sub
    '=============================================================================================
    'расстановка дорожных плит
    <CommandMethod("SetSlabAlignment")>
    Public Sub SetSlabAlignment()
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'ссылка на чертеж .dwg
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim db As Database = acDoc.Database
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
        Dim PutchTemplate As String = ""
        Dim NameFileTemplate As String = ""
        Dim arrayTemplates As String() = {}
        Dim countArrayPTemplates As Integer = 0
        Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
        Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
        Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName 'отделяем имя файла от папки
        Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName 'определяем имя родительской папки для папки где находится сборка
        If IO.Directory.Exists(parentPatch2 & "\FileResources\Sample\") Then
            PutchTemplate = parentPatch2 & "\FileResources\Sample\"
            For Each fileDwg As String In IO.Directory.GetFiles(PutchTemplate)
                Try
                    Dim File1 As String = IO.Path.GetFileName(fileDwg)
                    Dim ext As String = IO.Path.GetExtension(fileDwg)
                    If ext Like ".dwg" Then
                        ReDim Preserve arrayTemplates(countArrayPTemplates)
                        arrayTemplates(countArrayPTemplates) = Path.GetFileNameWithoutExtension(File1)
                        countArrayPTemplates += 1
                    End If
                Catch ex As ArgumentException
                End Try
            Next
        End If
        Dim formRoadSlab As FormInsertSlab = New FormInsertSlab
        formRoadSlab.ComboBox1.DataSource = arrayTemplates
        formRoadSlab.ShowDialog()
        If formRoadSlab.boolShow = False Then
            Exit Sub
        End If

        Dim templateXML As String = PutchTemplate & formRoadSlab.ComboBox1.Text & ".xml"
        Dim templateDwg As String = PutchTemplate & formRoadSlab.ComboBox1.Text & ".dwg"
        If File.Exists(templateXML) = False Then
            MsgBox("Шаблон в формате xml не найден. Макро прервано.")
            Exit Sub
        End If
        If File.Exists(templateXML) = False Then
            MsgBox("Шаблон в формате dwg не найден. Макро прервано.")
            Exit Sub
        End If
        Dim arrayPropertiesBLK As String(,) = {}
        Dim boolReadStyleDim As Boolean = TOPOSymbols.FuncReadPropertiesObjectToXML(templateXML, "Проектирование автомобильных дорог", "Временные дороги и проезды", "Блок", "Плита дорожная", arrayPropertiesBLK)
        Dim nameblk As String = "Плита дорожная (ПДН)"
        Dim nameLayer As String = "0"
        If arrayPropertiesBLK.Length > 0 Then
            nameblk = MathFunction.FuncFindValueToArray2d(arrayPropertiesBLK, "Name")
            Dim boolReadBlk As Boolean = AcadFunction.FuncCopyBlockDestDraw(templateDwg, nameblk)
            Dim tempNameLayer As String = MathFunction.FuncFindValueToArray2d(arrayPropertiesBLK, "Layer")
            Dim Idlayer As ObjectId = AcadFunction.FuncAddLayer(tempNameLayer, 7,, 20)
            If IsNothing(Idlayer) = False Then
                nameLayer = tempNameLayer
            End If
        End If
        Using acLckDoc As DocumentLock = acDoc.LockDocument()
            Dim tr As Transaction = db.TransactionManager.StartTransaction()
            Using tr
                Dim acBlkTbl As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                'загружаем блок
                Dim blk As BlockTableRecord = Nothing
                If acBlkTbl.Has(nameblk) = True Then
                    Dim blkID As ObjectId = acBlkTbl.Item(nameblk)
                    If blkID.IsNull = False Then
                        blk = tr.GetObject(blkID, OpenMode.ForRead)
                    End If
                End If
                If IsNothing(blk) = True Then
                    MsgBox("Не удалось найти блок <Плита жб>. Макро прервано.")
                    Exit Sub
                End If
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                '=====================================================================================================
                'выбираем трассу
                Dim Align As Autodesk.Civil.DatabaseServices.Alignment = Nothing
                Dim ColAlign As ObjectIdCollection = CivilDoc.GetAlignmentIds
                If ColAlign.Count = 0 Then
                    MsgBox("На чертеже нет объектов типа ТРАССА. Макро прервано.")
                    Exit Sub
                ElseIf ColAlign.Count = 1 Then
                    Align = tr.GetObject(ColAlign.Item(0), OpenMode.ForRead)
                Else
                    '=============================================================
                    'выбираем трассу
                    Dim ppr1 As PromptEntityResult = ed.GetEntity("Выберите трассу: ")
                    Dim ent1 As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(ppr1.ObjectId, OpenMode.ForRead)
                    If TypeOf ent1 Is Autodesk.Civil.DatabaseServices.Alignment Then
                        Align = ent1
                    End If
                End If
                If IsNothing(Align) = True Then
                    MsgBox("Ошибка выбора ТРАССЫ. Макро прервано.")
                    Exit Sub
                End If
                'преобразуем трассу в полилинию
                Dim polyAlignID As ObjectId = Align.GetPolyline
                If polyAlignID.IsNull = True Then
                    MsgBox("Не удалось преобразовать Трассу в полилинию. Макро прервано.")
                    Exit Sub
                End If
                Dim polyAlign As Polyline = polyAlignID.GetObject(OpenMode.ForWrite)
                Dim startPosition As Double = 0
                Dim ptSelect As PromptPointResult = ed.GetPoint("Укажите точку начала разбивки: ")
                If ptSelect.Status = PromptStatus.OK Then
                    Dim point As Point3d = ptSelect.Value
                    point = polyAlign.GetClosestPointTo(point, False)
                    startPosition = polyAlign.GetDistAtPoint(point)
                End If
                'делаем offset
                If formRoadSlab.NumericUpDown5.Value <> 0 Then
                    Dim dbPoly As DBObjectCollection = polyAlign.GetOffsetCurves(formRoadSlab.NumericUpDown5.Value)
                    If dbPoly.Count = 1 Then
                        Dim obj As Autodesk.AutoCAD.DatabaseServices.DBObject = dbPoly.Item(0)
                        If TypeOf obj Is Polyline Then
                            polyAlign = obj
                        End If
                    End If
                End If
                'начальные значения длины и ширины плиты
                Dim heightSlab As Double = formRoadSlab.NumericUpDown1.Value
                Dim widthSlab As Double = formRoadSlab.NumericUpDown2.Value
                Dim countLeftSlab As Integer = formRoadSlab.NumericUpDown3.Value
                Dim countRightSlab As Integer = formRoadSlab.NumericUpDown4.Value
                Dim numberSlab As Integer = formRoadSlab.NumericUpDown6.Value
                Dim mark As String = formRoadSlab.TextBox1.Text
                'Dim blkprevIns As BlockReference = AcadFunction.FuncInsertDrawingBlk(polyAlign.StartPoint, nameblk)
                'If IsNothing(blkprevIns) = False Then
                '    Dim boolAng As Boolean = AcadFunction.SetRotationParameter(blkprevIns, "Угол поворота", 0)
                '    Dim boolWidth As Boolean = AcadFunction.SetRotationParameter(blkprevIns, "Ширина плиты", heightSlab)
                '    Dim boolHeight As Boolean = AcadFunction.SetRotationParameter(blkprevIns, "Длина плиты", widthSlab)
                'End If
                Dim startPointPrefLeft As Point3d = New Point3d
                Dim endPointPrefLeft As Point3d = New Point3d
                Dim startPointPrefRight As Point3d = New Point3d
                Dim endPointPrefRight As Point3d = New Point3d
                Dim countSlab As Integer = 0

                For i As Double = startPosition To polyAlign.Length Step widthSlab
                    Try
                        Dim countErr As Integer = 0
                        If polyAlign.Length - i < widthSlab Then
                            Exit For
                        End If
line1:
                        If countErr > 5 Then
                            Exit For
                        End If
                        'точка вставки осевого блока
                        Dim ptInsert As Point3d = polyAlign.GetPointAtDist(startPosition)
                        'строим временную окружность с радиусом рамным длине плиты
                        Dim acCirc As Circle = New Circle()
                        acCirc.SetDatabaseDefaults()
                        acCirc.Center = ptInsert
                        acCirc.Radius = widthSlab
                        'ищем точку пересечения окружности и полилинии
                        Dim pt2Insert As Point3d = New Point3d(-1, -1, -1)
                        Dim pointCollection As Point3dCollection = New Point3dCollection()
                        polyAlign.IntersectWith(acCirc, Intersect.OnBothOperands, New Plane(), pointCollection, 0, 0)
                        If pointCollection.Count > 0 Then
                            For j As Integer = 0 To pointCollection.Count - 1
                                pt2Insert = pointCollection.Item(j)
                                Dim tempDist As Double = polyAlign.GetDistAtPoint(pt2Insert)
                                If tempDist > startPosition Then
                                    startPosition = tempDist
                                    Exit For
                                End If
                            Next j
                        End If
                        If pt2Insert.X = -1 And pt2Insert.Y = -1 And pt2Insert.Z = -1 Then
                            Exit For
                        End If
                        'угол разворота блока
                        Dim rotateBlk As Double = ptInsert.Convert2d(New Plane).GetVectorTo(pt2Insert.Convert2d(New Plane)).Angle
                        'углы перпендикулярные углу разворота блока
                        Dim rotBlk1 As Double = rotateBlk + Math.PI / 2
                        If rotBlk1 > Math.PI * 2 Then rotBlk1 -= Math.PI * 2
                        Dim rotBlk2 As Double = rotateBlk - Math.PI / 2
                        If rotBlk2 < 0 Then rotBlk2 += Math.PI * 2
                        'точки по верху блока
                        Dim newStartLeftPoint As Point3d = MathFunction.FuncReturnPolar3dCoordinates2(ptInsert.Convert2d(New Plane), rotBlk1, heightSlab * countLeftSlab)
                        Dim newStartRightPoint As Point3d = MathFunction.FuncReturnPolar3dCoordinates2(ptInsert.Convert2d(New Plane), rotBlk2, heightSlab * countRightSlab)
                        'точки по низу блока
                        Dim newEndLeftPoint As Point3d = MathFunction.FuncReturnPolar3dCoordinates2(pt2Insert.Convert2d(New Plane), rotBlk1, heightSlab * countLeftSlab)
                        Dim newEndtRightPoint As Point3d = MathFunction.FuncReturnPolar3dCoordinates2(pt2Insert.Convert2d(New Plane), rotBlk2, heightSlab * countRightSlab)
                        'проверяем пересечения отрезков если это не первый блок а последующие
                        If countSlab > 0 Then
                            'формируем боковые отрезки предыдущего блока
                            Dim linePrevLeft As Line = New Line(startPointPrefLeft, endPointPrefLeft)
                            Dim linePrevRight As Line = New Line(startPointPrefRight, endPointPrefRight)
                            Dim lineTop As Line = New Line(newStartLeftPoint, newStartRightPoint)
                            Dim lineBottom As Line = New Line(newEndLeftPoint, newEndtRightPoint)
                            Dim deltaL As Double = 0
                            'находим точку пересечения трассы и сечения
                            pt2Insert = New Point3d(-1, -1, -1)
                            pointCollection = New Point3dCollection()
                            linePrevLeft.IntersectWith(lineTop, Intersect.ExtendBoth, New Plane(), pointCollection, 0, 0)
                            If pointCollection.Count = 1 Then
                                For j As Integer = 0 To pointCollection.Count - 1
                                    pt2Insert = pointCollection.Item(j)
                                    Dim boolPoint As Boolean = MathFunction.IsPointOnLineSegment(linePrevLeft.StartPoint, linePrevLeft.EndPoint, pt2Insert)
                                    If boolPoint = True Then
                                        deltaL = endPointPrefLeft.Convert2d(New Plane).GetVectorTo(newStartLeftPoint.Convert2d(New Plane)).Length
                                    End If
                                Next j
                            End If
                            pointCollection = New Point3dCollection()
                            linePrevRight.IntersectWith(lineTop, Intersect.ExtendBoth, New Plane(), pointCollection, 0, 0)
                            If pointCollection.Count = 1 Then
                                For j As Integer = 0 To pointCollection.Count - 1
                                    pt2Insert = pointCollection.Item(j)
                                    Dim boolPoint As Boolean = MathFunction.IsPointOnLineSegment(linePrevRight.StartPoint, linePrevRight.EndPoint, pt2Insert)
                                    If boolPoint = True Then
                                        Dim tdeltaL As Double = endPointPrefRight.Convert2d(New Plane).GetVectorTo(newStartRightPoint.Convert2d(New Plane)).Length
                                        If tdeltaL > deltaL Then
                                            deltaL = tdeltaL
                                        End If
                                    End If
                                Next j
                            End If
                            'заново вычисляем позиции
                            If deltaL > 0.005 And countErr < 5 Then
                                countErr += 1
                                startPosition += deltaL
                                startPosition -= widthSlab
                                GoTo line1
                            End If
                        End If
                        'вставляем блоки. сначала центральный
                        Dim blkIns As BlockReference = AcadFunction.FuncInsertDrawingBlk(ptInsert, nameblk, nameLayer,,,, rotateBlk)
                        If IsNothing(blkIns) = False Then
                            Dim boolAng As Boolean = AcadFunction.SetRotationParameterSafe(blkIns, "Угол поворота", 0)
                            Dim boolWidth As Boolean = AcadFunction.SetRotationParameterSafe(blkIns, "Ширина плиты", heightSlab)
                            Dim boolHeight As Boolean = AcadFunction.SetRotationParameterSafe(blkIns, "Длина плиты", widthSlab)
                        End If
                        'вставляем блоки влево
                        For k As Integer = 1 To countLeftSlab - 1
                            Dim insPoint As Point3d = MathFunction.FuncReturnPolar3dCoordinates2(ptInsert.Convert2d(New Plane), rotBlk1, heightSlab * k + heightSlab)
                            blkIns = AcadFunction.FuncInsertDrawingBlk(insPoint, nameblk, nameLayer,,,, rotateBlk)
                            If IsNothing(blkIns) = False Then
                                Dim boolAng As Boolean = AcadFunction.SetRotationParameterSafe(blkIns, "Угол поворота", 0)
                                Dim boolWidth As Boolean = AcadFunction.SetRotationParameterSafe(blkIns, "Ширина плиты", heightSlab)
                                Dim boolHeight As Boolean = AcadFunction.SetRotationParameterSafe(blkIns, "Длина плиты", widthSlab)
                            End If
                        Next k
                        For k As Integer = 0 To countRightSlab - 1
                            Dim insPoint As Point3d = MathFunction.FuncReturnPolar3dCoordinates2(ptInsert.Convert2d(New Plane), rotBlk2, heightSlab * k + heightSlab)
                            blkIns = AcadFunction.FuncInsertDrawingBlk(insPoint, nameblk, nameLayer,,,, rotateBlk)
                            If IsNothing(blkIns) = False Then
                                Dim boolAng As Boolean = AcadFunction.SetRotationParameterSafe(blkIns, "Угол поворота", 0)
                                Dim boolWidth As Boolean = AcadFunction.SetRotationParameterSafe(blkIns, "Ширина плиты", heightSlab)
                                Dim boolHeight As Boolean = AcadFunction.SetRotationParameterSafe(blkIns, "Длина плиты", widthSlab)
                            End If
                        Next k
                        startPointPrefLeft = newStartLeftPoint
                        endPointPrefLeft = newEndLeftPoint
                        startPointPrefRight = newStartRightPoint
                        endPointPrefRight = newEndtRightPoint
                        countSlab += 1
                    Catch ex As Autodesk.AutoCAD.Runtime.Exception
                        MsgBox("Ошибка вставки блока в ряде " & countSlab)
                    End Try
                Next i
                tr.Commit()
            End Using
        End Using
        MsgBox("Макро завершено.")
    End Sub
    '=============================================================================================
    'Расстановка блоков вдоль полилинии
    <CommandMethod("SplitPolyline")>
    Public Sub SplitPolyline()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acDocEd As Editor = acDoc.Editor
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
        'Массив для вершин линий
        Dim ArrayPointVert As Double(,) = Nothing
        Dim CountArrayPointVert As Integer = 0
        Dim shagSect As Double = 3.5
        Dim pStrOpts As PromptDoubleOptions = New PromptDoubleOptions(vbLf & "Введите ширину секции ограждения: ")
        pStrOpts.DefaultValue = 3.5
        Dim pStrRes As PromptDoubleResult = acDoc.Editor.GetDouble(pStrOpts)
        shagSect = Val(pStrRes.Value)
        Dim count As Integer = 1
        Dim propBlk As String(,) = {}
        ReDim Preserve propBlk(1, 4)
        propBlk(0, 0) = "Name"
        propBlk(1, 0) = "A-Стойка металлическая сетчатого ограждения"

        Do
            Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
            Using tr
                Dim acBlkTbl As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord = tr.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                'выбираем полилинию
                '================================================================================================
                Dim resPolyline As PromptEntityResult = Nothing
                Dim opt1 As PromptEntityOptions = New PromptEntityOptions(vbLf & "Выберите полилинию: ")
                resPolyline = acDocEd.GetEntity(opt1)
                If Not (resPolyline.Status = PromptStatus.OK) Then
                    tr.Commit()
                    Exit Do
                End If
                Dim pvId1 As ObjectId = resPolyline.ObjectId
                Dim obj As Autodesk.AutoCAD.DatabaseServices.Entity = tr.GetObject(pvId1, OpenMode.ForWrite)
                If TypeOf obj Is Polyline Then
                    Dim poly As Polyline = obj
                    Dim endPos As Point3d = New Point3d
                    Dim rotationBlk As Double = 0
                    For i As Integer = 0 To poly.NumberOfVertices - 2
                        Dim userSeg As LineSegment3d = poly.GetLineSegmentAt(i)
                        Dim startDist As Double = poly.GetDistAtPoint(userSeg.StartPoint)
                        Dim endDist As Double = poly.GetDistAtPoint(userSeg.EndPoint)
                        rotationBlk = userSeg.StartPoint.Convert2d(New Plane).GetVectorTo(userSeg.EndPoint.Convert2d(New Plane)).Angle
                        Dim lSeg As Double = userSeg.StartPoint.GetVectorTo(userSeg.EndPoint).Length
                        Dim a As Integer = MathFunction.FuncTrimDigitFloo(lSeg / shagSect, 0)
                        Dim newLen As Double = shagSect * a
                        Dim deltaSeg As Double = lSeg - newLen
                        If userSeg.Length <= shagSect Then
                            propBlk(0, 1) = "Number"
                            propBlk(1, 1) = count
                            Dim idInsertBlk As BlockReference = AcadFunction.FuncInsBlkByArrayProperties(userSeg.StartPoint, propBlk, 0)
                            endPos = userSeg.StartPoint
                            count += 1
                        Else
                            For j As Double = startDist To endDist Step shagSect
                                propBlk(0, 1) = "Number"
                                propBlk(1, 1) = count
                                Dim insertPoint As Point3d = poly.GetPointAtDist(j)
                                Dim idInsertBlk As BlockReference = AcadFunction.FuncInsBlkByArrayProperties(insertPoint, propBlk, rotationBlk)
                                endPos = insertPoint
                                count += 1
                            Next j
                        End If
                    Next
                    If endPos.GetVectorTo(poly.EndPoint).Length > 0.5 Then
                        Dim idInsertBlk As BlockReference = AcadFunction.FuncInsBlkByArrayProperties(poly.EndPoint, propBlk, rotationBlk)
                        count += 1
                    End If
                End If
                tr.Commit()
            End Using
        Loop
    End Sub
End Class
