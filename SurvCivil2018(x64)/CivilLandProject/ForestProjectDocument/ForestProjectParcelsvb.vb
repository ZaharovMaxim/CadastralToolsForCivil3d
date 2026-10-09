Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports System.IO
Imports System
Imports Autodesk.Civil.ApplicationServices
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Windows
Imports CadastralTools
Imports NetTopologySuite
Public Class ForestProjectParcelsvb
    <CommandMethod("TD_CREATE")>
    Public Sub CreateTaxationDescription()
        Dim form As New TaxationDescriptionForm()
        Application.ShowModalDialog(form)
    End Sub
End Class
