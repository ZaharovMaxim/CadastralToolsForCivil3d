Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Windows
Public Class ProjectManagement
    '=============================================================================================
    'палитра с условными знаками (новая версия)
    <CommandMethod("ProjectLandManagement")>
    Public Sub ProjectLandManagement()
        '=========================================================================================
        'инициализируем формы
        Dim PaletteLandProject As TerritoryPlanProject = New TerritoryPlanProject
        Dim landProject As PaletteSet = New Autodesk.AutoCAD.Windows.PaletteSet("Диспетчер проектов")
        '========================================================================================
        Dim UserSize As Drawing.Size = New Drawing.Size(350, 620)
        'landProject.DockEnabled = DockSides.Left + DockSides.Right
        landProject.AddVisual("Диспетчер проектов", PaletteLandProject)
        landProject.Size = UserSize
        landProject.Visible = True
    End Sub
End Class
