Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Autodesk.Gis.Map
Imports Autodesk.Gis.Map.Constants
Imports Autodesk.Gis.Map.Query
Imports Microsoft.Office.Interop
Imports OSGeo.MapGuide
Imports System.IO
Imports System.Windows.Forms
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application

Public Class TransformDwgToMap3d
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    <CommandMethod("DwgObjectTransform")>
    Public Sub DwgObjectTransform()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim ed As Editor = acDoc.Editor
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
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'проверяем систему координат текущего чертежа
        Dim CoordSystemCode As String = CivilApplication.ActiveDocument.Settings.DrawingSettings.UnitZoneSettings.CoordinateSystemCode
        Dim coordSysFactory As MgCoordinateSystemFactory = New MgCoordinateSystemFactory()
        Dim csCatalog As MgCoordinateSystemCatalog = coordSysFactory.GetCatalog()
        Dim csDict As MgCoordinateSystemDictionary = csCatalog.GetCoordinateSystemDictionary()
        Dim DrawingSysCoord As MgCoordinateSystem = csDict.GetCoordinateSystem(CoordSystemCode)
        If DrawingSysCoord.CsCode Like "" Then
            MsgBox("Система координат не присвоена текущему чертежу!!! Макро прервано")
            Exit Sub
        End If
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        'загружаем целевой чертеж
        Dim fileDialog1 As Autodesk.AutoCAD.Windows.OpenFileDialog = New Autodesk.AutoCAD.Windows.OpenFileDialog("Открыть файл dwg", "", "dwg", "Открыть файл-источник", Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.AllowAnyExtension)
        Dim FileRez As String = fileDialog1.ShowDialog()
        Dim PutchDWG As String = fileDialog1.Filename
        If IsNothing(PutchDWG) = True Then Exit Sub
        If PutchDWG.Trim.Length = 0 Then Exit Sub
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Dim sourceDb As Database = New Database(False, True)
        Dim destCivilDoc As CivilDocument = Autodesk.Civil.ApplicationServices.CivilDocument.GetCivilDocument(sourceDb)
        sourceDb.ReadDwgFile(PutchDWG, FileShare.Read, True, String.Empty)
        Dim sourceTr As Transaction = sourceDb.TransactionManager.StartTransaction()
        Using sourceTr
            If destCivilDoc.Settings.DrawingSettings.UnitZoneSettings.CoordinateSystemCode.Trim Like "." Then
                MsgBox("Система координат не присвоена выбранному чертежу!!! Макро прервано")
                Exit Sub
            End If
        End Using
        sourceDb.CloseInput(True)
        Dim docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Try
                Dim mapApp As MapApplication = HostMapApplicationServices.Application
                Dim activeProject As Project.ProjectModel = mapApp.ActiveProject
                Dim PROJECT As Project.ProjectModel = HostMapApplicationServices.Application.ActiveProject
                Dim AttchDrawing As Project.AttachedDrawing = PROJECT.DrawingSet.AttachDrawing(PutchDWG)
                AttchDrawing.Activate()
                Dim qryModel As QueryModel = PROJECT.CreateQuery()
                qryModel.Clear()
                Dim locationCondition As Query.LocationCondition = New Query.LocationCondition
                Dim qryCondition As LocationCondition = New LocationCondition()
                qryCondition.Boundary = AllBoundary.Create
                Dim qryBranch As New QueryBranch(JoinOperator.OperatorAnd)
                qryBranch.AppendOperand(qryCondition)
                qryModel.Define(qryBranch)
                qryModel.Mode = QueryType.QueryDraw
                qryModel.Run()
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                AttchDrawing.Deactivate()
                PROJECT.DrawingSet.DetachDrawing(AttchDrawing.ObjectId)
            Catch ex As Autodesk.Gis.Map.MapException
                MsgBox("Не удалось трансформировать чертеж. Попробуйте создать псевдоним пути к Вашему файлу.")
            End Try
            acTrans.Commit()
        End Using
        docLock.Dispose()
    End Sub
End Class
