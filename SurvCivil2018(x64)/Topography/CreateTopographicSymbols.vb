Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Windows
Public Class CreateTopographicSymbols
    '=============================================================================================
    'палитра с условными знаками (новая версия)
    <CommandMethod("ProjectCadastralSumbols")>
    Public Sub CreateTopographicSymbols()
        '=========================================================================================
        'инициализируем формы
        Dim PaletteTopoSumbol As TOPOSymbols = New TOPOSymbols
        Dim PaletteProperties As UserPropertiesTopoSymbols = New UserPropertiesTopoSymbols
        PaletteTopoSumbol.UserProperties = PaletteProperties
        Dim TOPO As PaletteSet = New Autodesk.AutoCAD.Windows.PaletteSet("Условные знаки")
        '========================================================================================
        Dim UserSize As Drawing.Size = New Drawing.Size(300, 800)
        TOPO.Size = UserSize
        TOPO.DockEnabled = DockSides.Left + DockSides.Right
        TOPO.AddVisual("Условные знаки", PaletteTopoSumbol)
        TOPO.AddVisual("Установки", PaletteProperties)
        TOPO.Visible = True
    End Sub
End Class
