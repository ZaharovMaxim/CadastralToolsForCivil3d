Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports System.Xml
Imports Autodesk.AutoCAD.Runtime
Public Class CreateProject
    <CommandMethod("CreateLandProjectPlan")>
    Public Sub CreateLandProjectPlan()
        Dim formCreateProject As FormCreateProject = New FormCreateProject
        formCreateProject.ShowDialog()

    End Sub
End Class
