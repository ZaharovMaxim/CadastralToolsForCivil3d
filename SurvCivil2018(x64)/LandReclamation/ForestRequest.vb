Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.Civil.DatabaseServices
Imports Microsoft.Office.Interop
Imports OSGeo.MapGuide
Imports System.IO
Imports System.Windows.Forms
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
Imports NetTopologySuite.Geometries
Imports NetTopologySuite.IO
Imports NetTopologySuite.Operation.Buffer
Imports System.Xml
Public Class ForestRequest
    <CommandMethod("ForestRequestXml")>
    Public Sub ForestRequestXml()
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        Dim schemasDir As String = "k:\Civil3DToolsUtility\NET\Schemas\"
        Dim outPath As String = "D:\Temp\out.xml"

        ' === 1. Показываем форму ===
        Using form As New ForestRequestForm()
            If form.ShowDialog() <> DialogResult.OK Then
                ed.WriteMessage("Отменено пользователем." & vbLf)
                Return
            End If

            ' === 2. Собираем XML из данных формы ===
            ed.WriteMessage("=== Формирование XML ===" & vbLf)
            Dim doc = ExtractRequestBuilder.Build(form.ResultData)
            XmlHelper.Save(doc, outPath)
            ed.WriteMessage($"XML сохранён: {outPath}" & vbLf)

            ' === 3. Загружаем XSD ===
            ed.WriteMessage("=== Загрузка XSD ===" & vbLf)
            Dim schemas = XmlValidator.LoadSchemasForChangeRequest(schemasDir)

            ' === 4. Валидируем ===
            ed.WriteMessage("=== Валидация ===" & vbLf)
            Dim errors = XmlValidator.Validate(outPath, schemas)

            If errors.Count = 0 Then
                ed.WriteMessage("OK: XML валиден." & vbLf)
            Else
                ed.WriteMessage($"Найдено ошибок: {errors.Count}" & vbLf)
                For Each userErr In errors
                    ed.WriteMessage("  " & userErr.ToString() & vbLf)
                Next
            End If
        End Using
    End Sub
    <CommandMethod("EditTaxation")>
    Public Sub EditTaxationCommand()
        Using f As New EditTaxationForm()
            Application.ShowModalDialog(f)
        End Using
    End Sub
End Class
