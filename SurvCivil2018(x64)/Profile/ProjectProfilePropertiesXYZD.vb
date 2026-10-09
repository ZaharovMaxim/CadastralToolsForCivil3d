Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports System.Configuration
Imports Microsoft.Office.Interop
Imports Autodesk.AutoCAD.Colors
Public Class ProjectProfilePropertiesXYZD
    <CommandMethod("ProjProfileXYZD")>
    Public Sub ProjProfileXYZD()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim docCol As Autodesk.AutoCAD.ApplicationServices.DocumentCollection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager
        Dim CivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilApplication.ActiveDocument()
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
        Dim DocColectionAlignment As ObjectIdCollection = CivilDoc.GetAlignmentIds
        If DocColectionAlignment.Count = 0 Then
            MsgBox("В данном чертеже нет ТРАСС!!! Макро прервано")
            Exit Sub
        End If
        Dim locker As Autodesk.AutoCAD.ApplicationServices.DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument
        Dim tr As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction()
        'запуск диалогового окна
        Dim FormProjProfileXYZD As ProjProfilePropXYZD = New ProjProfilePropXYZD
        FormProjProfileXYZD.ShowDialog()
        If FormProjProfileXYZD.boolProf = False Then Exit Sub
        If IsNothing(FormProjProfileXYZD.AlignLine) = True Then
            tr.Commit()
            locker.Dispose()
            MsgBox("Трасса не выбрана! Макро прервано.")
            Exit Sub
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim ArrayXYZ As String(,) = Nothing 'массив для записи в Excel (0-пикет, 1-северное положение, 2-восточное положение, 3-отметка, 4-описание)
        Dim Count As Integer = 0 'счетчик элементов массива
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'точность координат и пикетов
        'устанавливаем системную переменную указаннаю пользователем
        Dim boolSist As Boolean = False
        'смотрим точность системной переменной до запуска программы и сохраняем ее
        Dim SistPerem As Integer = CivilDoc.Settings.DrawingSettings.AmbientSettings.Station.Precision.Value
        Dim RoundZn As Integer = FormProjProfileXYZD.NumericUpDown2.Value
        If SistPerem <> RoundZn Then 'устанавливаем новое значение
            CivilDoc.Settings.DrawingSettings.AmbientSettings.Station.Precision.Value = RoundZn
            boolSist = True
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Using tr
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'ссылка на трассу
            Dim align As Alignment = Nothing
            align = FormProjProfileXYZD.AlignLine
            If IsNothing(align) = True Then Exit Sub
            '========================================================================
            'ссылка на профиль
            Dim Prof As Autodesk.Civil.DatabaseServices.Profile = Nothing
            Prof = FormProjProfileXYZD.ProfLine
            If IsNothing(Prof) = True Then Exit Sub
            'значения начального и конечного пикетов
            Dim PKStartAlign As Double = FormProjProfileXYZD.ProfStart
            Dim PKEndAlign As Double = FormProjProfileXYZD.ProfEnd
            'шаг разбивки
            Dim Chag As Double = FormProjProfileXYZD.NumericUpDown1.Value
            '========================================================================================================
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'находим массив с рубленными пикетами
            Dim ArrayStEqw As Double(,) = Nothing
            Dim CountStEqw As Integer = 0
            Dim stationEqw As StationEquationCollection = align.StationEquations
            If stationEqw.Count > 0 Then
                For i As Integer = 0 To stationEqw.Count - 1
                    Dim St As StationEquation = stationEqw.Item(i)
                    ReDim Preserve ArrayStEqw(3, i)
                    ArrayStEqw(0, i) = St.RawStationBack
                    ArrayStEqw(1, i) = St.StationBack
                    ArrayStEqw(2, i) = St.StationAhead
                    ArrayStEqw(3, i) = 0
                Next i
            End If
            '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            'не учитывать горизонтальную геометрию
            If FormProjProfileXYZD.RadioButton2.Checked = True Then
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                Dim ArrayObject As Double() = {PKStartAlign, PKEndAlign, 3}
                Dim boolFuncAlignObject As Boolean = FuncAlignment.FuncReadAlignObjectXYZD(align, Prof, ArrayObject, ArrayXYZ, Chag, RoundZn, ArrayStEqw)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'горизонтальная геометрия учитывается
            ElseIf FormProjProfileXYZD.RadioButton1.Checked = True Then
                '============================================================
                'перебираем все элементы трассы для их сортировки в процессе возрастания
                Dim ArrayEntity As Double(,) = Nothing
                For Each myAe As AlignmentEntity In align.Entities
                    Dim AlignEnt As AlignmentEntityType = myAe.EntityType
                    Select Case AlignEnt
                    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                    'прямой участок
                        Case AlignmentEntityType.Line 'прямой участок
                            Dim AlignLine As AlignmentLine = myAe
                            Dim boolLineChar As Boolean = FuncAlignment.FuncCharacteristicsLine(align, AlignLine, ArrayEntity, PKStartAlign, PKEndAlign)
     '===========================================================================================================================
'дуга одиночная
                        Case AlignmentEntityType.Arc
                            Dim AlignArc As AlignmentArc = myAe
                            Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsArc(align, AlignArc, ArrayEntity, PKStartAlign, PKEndAlign)
'========================================================================================================================================
'спираль одиночная
                        Case AlignmentEntityType.Spiral
                            If TypeOf myAe Is AlignmentSTS Then
                                Dim AlignSTS As AlignmentSTS = myAe 'составная спираль спираль дуга спираль
                                Dim AlignS1 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    AlignS1 = AlignSTS.SpiralIn
                                Catch ex As System.InvalidOperationException
                                End Try
                                Try
                                    AlignS1 = AlignSTS.SpiralOut
                                Catch ex As System.InvalidOperationException
                                End Try
                                'если спираль существует
                                If IsNothing(AlignS1) = False Then
                                    Dim boolSpcChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, AlignS1, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                '==================================================================================================================================
                            ElseIf TypeOf myAe Is AlignmentSCS Then
                                Dim AlignSCS As AlignmentSCS = myAe 'составная спираль спираль прямая спираль
                                Dim AlignS2 As AlignmentSubEntitySpiral = Nothing
                                Try
                                    AlignS2 = AlignSCS.SpiralIn
                                Catch ex As System.InvalidOperationException
                                End Try
                                Try
                                    AlignS2 = AlignSCS.SpiralOut
                                Catch ex As System.InvalidOperationException
                                End Try
                                'если спираль существует
                                If IsNothing(AlignS2) = False Then
                                    Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, AlignS2, ArrayEntity, PKStartAlign, PKEndAlign)
                                End If
                                '=================================================================================================
                            ElseIf TypeOf myAe Is AlignmentSpiral Then
                                Dim AlignS3 As AlignmentSpiral = Nothing 'отдельная спираль
                                Try
                                    AlignS3 = myAe
                                Catch ex As System.InvalidOperationException
                                End Try
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSpiral(align, AlignS3, ArrayEntity, PKStartAlign, PKEndAlign)
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
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, Spin, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc1 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc1 = AlignSCS.Arc
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc1) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(align, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim SpOut As AlignmentSubEntitySpiral = Nothing
                            Try
                                SpOut = AlignSCS.SpiralOut
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(SpOut) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, SpOut, ArrayEntity, PKStartAlign, PKEndAlign)
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
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, Spin, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Line1 As AlignmentSubEntityLine = Nothing
                            Try
                                Line1 = AlignSLS.Tangent
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Line1) = False Then
                                Dim boolLineChar As Boolean = FuncAlignment.FuncCharacteristicsSubLine(align, Line1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim SpOut As AlignmentSubEntitySpiral = Nothing
                            Try
                                SpOut = AlignSLS.SpiralOut
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(SpOut) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, SpOut, ArrayEntity, PKStartAlign, PKEndAlign)
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
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, Spin, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Line1 As AlignmentSubEntityLine = Nothing
                            Try
                                Line1 = AlignSL.Tangent
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Line1) = False Then
                                Dim boolLineChar As Boolean = FuncAlignment.FuncCharacteristicsSubLine(align, Line1, ArrayEntity, PKStartAlign, PKEndAlign)
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
                                Dim boolLineChar As Boolean = FuncAlignment.FuncCharacteristicsSubLine(align, Line1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim SpOut As AlignmentSubEntitySpiral = Nothing
                            Try
                                SpOut = AlignLS.SpiralOut
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(SpOut) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, SpOut, ArrayEntity, PKStartAlign, PKEndAlign)
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
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, Spin, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc1 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc1 = AlignSC.Arc
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc1) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(align, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
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
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(align, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim SpOut As AlignmentSubEntitySpiral = AlignCS.SpiralOut
                            Try
                                SpOut = AlignCS.SpiralOut
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(SpOut) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, SpOut, ArrayEntity, PKStartAlign, PKEndAlign)
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
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, Sp1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Sp2 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp2 = AlignSSCSS.Spiral2
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp2) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, Sp2, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc1 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc1 = AlignSSCSS.Arc
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc1) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(align, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Sp3 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp3 = AlignSSCSS.Spiral3
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp3) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, Sp3, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Sp4 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp4 = AlignSSCSS.Spiral4
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp4) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, Sp4, ArrayEntity, PKStartAlign, PKEndAlign)
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
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(align, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc2 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc2 = AlignCCRC.Arc2
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc2) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(align, Arc2, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc3 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc3 = AlignCCRC.Arc3
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc3) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(align, Arc3, ArrayEntity, PKStartAlign, PKEndAlign)
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
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(align, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc2 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc2 = AlignCRC.Arc2
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc2) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(align, Arc2, ArrayEntity, PKStartAlign, PKEndAlign)
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
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(align, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Sp3 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp3 = AlignCSS.Spiral3
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp3) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, Sp3, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Sp4 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp4 = AlignCSS.Spiral4
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp4) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, Sp4, ArrayEntity, PKStartAlign, PKEndAlign)
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
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, Sp1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            If IsNothing(Sp2) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, Sp2, ArrayEntity, PKStartAlign, PKEndAlign)
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
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, Sp1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Sp2 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp2 = AlignSSC.Spiral2
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp2) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, Sp2, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc1 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc1 = AlignSSC.Arc
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc1) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(align, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
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
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, Sp1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc1 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc1 = AlignSCSSCS.Arc1
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc1) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(align, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Sp2 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp2 = AlignSCSSCS.Spiral2
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp2) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, Sp2, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Sp3 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp3 = AlignSCSSCS.Spiral3
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp3) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, Sp3, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc2 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc2 = AlignSCSSCS.Arc2
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc2) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(align, Arc2, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Sp4 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp4 = AlignSCSSCS.Spiral4
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp4) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, Sp4, ArrayEntity, PKStartAlign, PKEndAlign)
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
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, Sp1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc1 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc1 = AlignSCSCS.Arc1
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc1) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(align, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Sp2 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp2 = AlignSCSCS.Spiral2
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp2) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, Sp2, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc2 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc2 = AlignSCSCS.Arc2
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc2) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(align, Arc2, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Sp3 As AlignmentSubEntitySpiral = Nothing
                            Try
                                Sp3 = AlignSCSCS.Spiral3
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Sp3) = False Then
                                Dim boolSpChar As Boolean = FuncAlignment.FuncCharacteristicsSubSpiral(align, Sp3, ArrayEntity, PKStartAlign, PKEndAlign)
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
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(align, Arc1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Line1 As AlignmentSubEntityLine = Nothing
                            Try
                                Line1 = AlignCLC.Tangent
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Line1) = False Then
                                Dim boolLineChar As Boolean = FuncAlignment.FuncCharacteristicsSubLine(align, Line1, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                            Dim Arc2 As AlignmentSubEntityArc = Nothing
                            Try
                                Arc2 = AlignCLC.Arc1
                            Catch ex As System.InvalidOperationException
                            End Try
                            If IsNothing(Arc2) = False Then
                                Dim boolArcChar As Boolean = FuncAlignment.FuncCharacteristicsSubArc(align, Arc2, ArrayEntity, PKStartAlign, PKEndAlign)
                            End If
                    End Select
                Next
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'второй этап
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'сортируем массив
                If IsArray(ArrayEntity) = True Then
                    If ArrayEntity.GetUpperBound(1) > 0 Then
                        Dim boolArray As Boolean = MathFunction.FuncSortDblArray(ArrayEntity, 0)
                    End If
                Else
                    MsgBox("Не удалось отсортировать массив. Макро прервано!!!")
                    Exit Sub
                End If
                '=============================================================================================================
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                'начинаем перебирать полученный массив
                Dim CountArrayXYZ As Integer = 0
                For i As Integer = 0 To ArrayEntity.GetUpperBound(1)
                    'формируем массив с пикетами подлежащими записи
                    Dim ArrayObject As Double(,) = Nothing
                    ReDim Preserve ArrayObject(2, 1)
                    If ArrayEntity(4, i) = 0 Then 'прямой участок
                        ArrayObject(0, 0) = ArrayEntity(0, i)
                        ArrayObject(1, 0) = 0
                        ArrayObject(2, 0) = 0
                        ArrayObject(0, 1) = ArrayEntity(1, i)
                        ArrayObject(1, 1) = 0
                        ArrayObject(2, 1) = 0
                    ElseIf ArrayEntity(4, i) = 1 Then 'дуга
                        ArrayObject(0, 0) = ArrayEntity(0, i)
                        ArrayObject(1, 0) = 2
                        ArrayObject(2, 0) = 0
                        ArrayObject(0, 1) = ArrayEntity(1, i)
                        ArrayObject(1, 1) = 3
                        ArrayObject(2, 1) = 0
                    ElseIf ArrayEntity(4, i) = 2 Then
                        ArrayObject(0, 0) = ArrayEntity(0, i)
                        ArrayObject(1, 0) = 1
                        ArrayObject(2, 0) = 0
                        ArrayObject(0, 1) = ArrayEntity(1, i)
                        ArrayObject(1, 1) = 4
                        ArrayObject(2, 1) = 0
                    End If
                    '=======================================================================================
                    'пишем в этот массив рубленные пикеты
                    Dim CountArrayObject As Integer = ArrayObject.GetUpperBound(1) + 1
                    Dim CollSqPiket As StationEquationCollection = align.StationEquations
                    If CollSqPiket.Count > 0 Then
                        For Each STeq As StationEquation In CollSqPiket
                            Dim SqPK As Double = STeq.RawStationBack
                            If SqPK >= ArrayEntity(0, i) And SqPK < ArrayEntity(1, i) Then
                                ReDim Preserve ArrayObject(2, CountArrayObject)
                                ArrayObject(0, CountArrayObject) = SqPK
                                ArrayObject(1, CountArrayObject) = 6
                                ArrayObject(2, CountArrayObject) = 0
                                CountArrayObject += 1
                            End If
                        Next
                    End If
                    '====================================================================================================
                    'находим все точки с определенным шагом
                    'Dim ArrayObject As Double() = {ArrayEntity(0, i), ArrayEntity(1, i), ArrayEntity(4, i)}
                    Dim TempArrayXYZ As Double(,) = Nothing
                    'Dim boolFuncAlignObject As Boolean = FuncAlignObjectXYZ2(align, Prof, ArrayObject, TempArrayXYZ, Chag, RoundZn, ArrayStEqw)
                    Dim boolFuncAlignObject As Boolean = FuncAlignment.FuncSplitAlignObjectStationXYHD(align, Prof, ArrayObject, TempArrayXYZ, Chag)
                    If IsArray(TempArrayXYZ) = True Then
                        'копируем значения в сводный массив
                        For j As Integer = 0 To TempArrayXYZ.GetUpperBound(1)
                            Dim PK As Double = TempArrayXYZ(0, j)
                            'начинаем поиск в массиве
                            Dim boolCopy As Boolean = True
                            If IsArray(ArrayXYZ) = True Then 'если массив существует
                                For k As Integer = 0 To ArrayXYZ.GetUpperBound(1)
                                    Dim PKTemp As Double = ArrayXYZ(0, k)
                                    If Math.Abs((PKTemp - PK)) < 0.01 Then
                                        Dim DeskIsh As Double = ArrayXYZ(4, k)
                                        Dim Desk As Double = TempArrayXYZ(4, j)
                                        If Desk > 0 And DeskIsh = 0 Then
                                            ArrayXYZ(4, k) = Desk
                                        ElseIf DeskIsh = 3 Or DeskIsh = 4 Then
                                            If Desk = 1 Or Desk = 2 Then
                                                ArrayXYZ(4, k) = Desk
                                            End If
                                        End If
                                        boolCopy = False
                                    End If
                                Next k
                            End If
                            If boolCopy = True Then
                                ReDim Preserve ArrayXYZ(4, CountArrayXYZ)
                                ArrayXYZ(0, CountArrayXYZ) = TempArrayXYZ(0, j)
                                ArrayXYZ(1, CountArrayXYZ) = TempArrayXYZ(1, j)
                                ArrayXYZ(2, CountArrayXYZ) = TempArrayXYZ(2, j)
                                ArrayXYZ(3, CountArrayXYZ) = TempArrayXYZ(3, j)
                                ArrayXYZ(4, CountArrayXYZ) = TempArrayXYZ(4, j)
                                CountArrayXYZ += 1
                            End If
                        Next j
                    End If
                Next i
            End If
            'записываем полученный массив с данными в эксель
            'запускаем Excel
            If FormProjProfileXYZD.RadioButton3.Checked = True Then
                'находим все директории с файлами
                Dim myPath As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly() 'патч к сборке
                Dim PathFile As String = myPath.Location 'строка путь к файлу со сборкой
                'отделяем имя файла от папки
                Dim parentPatch1 As String = IO.Directory.GetParent(PathFile).FullName
                'определяем имя родительской папки для папки где находится сборка
                Dim parentPatch2 As String = IO.Directory.GetParent(parentPatch1).FullName
                '3. путь к таблицам
                Dim PutchTables As String = FormProjProfileXYZD.ComboBox3.Tag & FormProjProfileXYZD.ComboBox3.Text 'пути для шаблона 

                If IO.File.Exists(PutchTables) Then
                    Dim xlApp As Excel.Application = Nothing
                    Dim xlBook As Excel.Workbook = Nothing
                    Dim xlSheets As Excel.Worksheet = Nothing
                    Dim cell1 As Excel.Range 'первая ячейка в строке
                    Dim cell2 As Excel.Range 'последняя ячейка в строке
                    Dim CellRng As Excel.Range 'массив ячеек
                    xlApp = CreateObject("Excel.Application")
                    xlBook = xlApp.Workbooks.Open(PutchTables)
                    xlApp.Visible = True
                    'находим нужный лист(Ведомость объемов работ)
                    Dim SheetCountZu1 As Integer = 1
                    Dim xlsheet As Excel.Worksheet
                    For Each xlsheet In xlBook.Worksheets
                        If xlsheet.Name.ToString.Trim Like "Ведомость рабочих отметок" Then
                            SheetCountZu1 = xlsheet.Index
                            Exit For
                        End If
                    Next
                    xlsheet = xlBook.Worksheets.Item(SheetCountZu1) 'делаем ссылку на этот лист
                    xlsheet.Activate()
                    xlsheet.Cells(2, 1) = "Наименование трассы: " & FormProjProfileXYZD.AlignLine.Name
                    xlsheet.Cells(3, 1) = "Наименование профиля: " & FormProjProfileXYZD.ProfLine.Name
                    xlsheet.Cells(4, 1) = "Начальный ПК: " & FormProjProfileXYZD.TextBox1.Text
                    xlsheet.Cells(5, 1) = "Конечный ПК: " & FormProjProfileXYZD.TextBox2.Text
                    Dim exlPos As Integer = 8
                    If IsArray(ArrayXYZ) = True Then
                        For i As Integer = 0 To ArrayXYZ.GetUpperBound(1)
                            cell1 = xlsheet.Cells(exlPos, 1) 'ссылка на первую ячейку
                            cell2 = xlsheet.Cells(exlPos, 5) 'ссылка на 6 ячейку в строке
                            CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                            CellRng.EntireRow.Insert() 'вставляем строку
                            Dim NumbFormat As String = "0.00"
                            If RoundZn = 0 Then
                                NumbFormat = "0"
                            ElseIf RoundZn = 1 Then
                                NumbFormat = "0.0"
                            ElseIf RoundZn = 2 Then
                                NumbFormat = "0.00"
                            ElseIf RoundZn = 3 Then
                                NumbFormat = "0.000"
                            ElseIf RoundZn = 4 Then
                                NumbFormat = "0.00000"
                            End If
                            cell1 = xlsheet.Cells(exlPos, 2) 'ссылка на первую ячейку
                            cell2 = xlsheet.Cells(exlPos, 4) 'ссылка на 6 ячейку в строке
                            CellRng = xlsheet.Range(cell1, cell2) 'выбираем строку
                            CellRng.NumberFormat = NumbFormat
                            Dim PKStrFormat As Double = ArrayXYZ(0, i)
                            'переводим все в форматированный пикет
                            Dim PKSTRFormat2 As String = FuncAlignment.FuncFormatRublPiket(align, PKStrFormat, RoundZn, 0)
                            Dim x As Double = Val(ArrayXYZ(1, i))
                            Dim y As Double = Val(ArrayXYZ(2, i))
                            Dim h As Double = Val(ArrayXYZ(3, i))
                            Dim Note As Double = ArrayXYZ(4, i)
                            Dim NoteStr As String = "'-"
                            If Note = 0 Then
                                NoteStr = "'-"
                            ElseIf Note = 1 Then
                                NoteStr = "НПК"
                            ElseIf Note = 2 Then
                                NoteStr = "НКК"
                            ElseIf Note = 3 Then
                                NoteStr = "ККК"
                            ElseIf Note = 4 Then
                                NoteStr = "КПК"
                            ElseIf Note = 5 Then
                                NoteStr = "СК"
                            ElseIf Note = 6 Then
                                NoteStr = "Руб.ПК"
                            End If
                            xlsheet.Cells(exlPos, 1) = PKSTRFormat2
                            xlsheet.Cells(exlPos, 2) = y
                            xlsheet.Cells(exlPos, 3) = x
                            xlsheet.Cells(exlPos, 4) = h
                            xlsheet.Cells(exlPos, 5) = NoteStr
                            exlPos += 1
                        Next i
                        xlSheets = Nothing
                        xlBook = Nothing
                        xlApp.Quit()
                        xlApp = Nothing
                    End If
                End If
            ElseIf FormProjProfileXYZD.RadioButton4.Checked = True Then
                '===============================================================================================
                'вставляем слой для таблицы
                'вставляем слой для тангенциального хода и номеров вершин углов
                Dim NameLayer As String = "1-Ведомости"
                Dim IDLayerIns As ObjectId = AcadFunction.FuncAddLayer(NameLayer, 7,, 20)
                Dim bt As BlockTable = tr.GetObject(acCurDb.BlockTableId, OpenMode.ForWrite)
                Dim ms As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim pPtRes As PromptPointResult
                Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
                pPtOpts.Message = vbLf & "Укажите точку вставки таблицы: "
                pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                Dim ptStart As Point3d = pPtRes.Value
                If pPtRes.Status = PromptStatus.Cancel Then GoTo LineEnd
                Dim tb As Autodesk.AutoCAD.DatabaseServices.Table = New Autodesk.AutoCAD.DatabaseServices.Table()
                tb.TableStyle = acCurDb.Tablestyle
                tb.SetDatabaseDefaults()
                ms.UpgradeOpen()
                ms.AppendEntity(tb)
                tb.Position = ptStart
                If IsNothing(IDLayerIns) = False Then
                    If IDLayerIns.IsNull = False Then
                        tb.LayerId = IDLayerIns
                    End If
                End If
                tb.SetSize(3, 5)
                tb.SetRowHeight(5)
                tb.Columns(0).Width = 20
                tb.Columns(1).Width = 25
                tb.Columns(2).Width = 25
                tb.Columns(3).Width = 15
                tb.Columns(4).Width = 15
                tb.Cells(0, 0).Value = "Ведомость рабочих отметок продольного профиля"
                tb.Cells(1, 0).Value = "Пикет+"
                tb.Cells(1, 1).Value = "X м."
                tb.Cells(1, 2).Value = "Y м."
                tb.Cells(1, 3).Value = "H м."
                tb.Cells(1, 4).Value = "Примечание"
                Dim NumRows As Integer = 2
                If IsArray(ArrayXYZ) = True Then
                    For i As Integer = 0 To ArrayXYZ.GetUpperBound(1)
                        tb.InsertRows(NumRows, 2.5, 1)
                        Dim PKStrFormat As Double = ArrayXYZ(0, i)
                        'переводим все в форматированный пикет
                        Dim PKSTRFormat2 As String = FuncAlignment.FuncFormatRublPiket(align, PKStrFormat, RoundZn, 0)
                        Dim x As Double = Val(ArrayXYZ(1, i))
                        Dim y As Double = Val(ArrayXYZ(2, i))
                        Dim h As Double = Val(ArrayXYZ(3, i))
                        Dim Note As Double = ArrayXYZ(4, i)
                        Dim NoteStr As String = "-"
                        If Note = 0 Then
                            NoteStr = "-"
                        ElseIf Note = 1 Then
                            NoteStr = "НПК"
                        ElseIf Note = 2 Then
                            NoteStr = "НКК"
                        ElseIf Note = 3 Then
                            NoteStr = "ККК"
                        ElseIf Note = 4 Then
                            NoteStr = "КПК"
                        ElseIf Note = 5 Then
                            NoteStr = "СК"
                        ElseIf Note = 6 Then
                            NoteStr = "Руб.ПК"
                        End If
                        tb.Cells(NumRows, 0).Value = PKSTRFormat2
                        tb.Cells(NumRows, 1).Value = FuncFormatZN.FuncFormatCoordinate(y, RoundZn)
                        tb.Cells(NumRows, 2).Value = FuncFormatZN.FuncFormatCoordinate(x, RoundZn)
                        tb.Cells(NumRows, 3).Value = FuncFormatZN.FuncFormatCoordinate(h, RoundZn)
                        tb.Cells(NumRows, 4).Value = NoteStr
                        NumRows += 1
                    Next i
                End If
                tr.AddNewlyCreatedDBObject(tb, True)
            End If
LineEnd:
            locker.Dispose()
            tr.Commit()
        End Using
    End Sub
End Class
