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
Public Class TransformationCoordinates
    <CommandMethod("CalcParametersAffineTransform")>
    Public Sub CalcParametersAffineTransform()
        'создаем класс с новой формой
        Dim Transform1 As TransformCoordinates = New TransformCoordinates
        Transform1.ShowDialog()
        If Transform1.stbool = False Then Exit Sub
    End Sub
End Class
