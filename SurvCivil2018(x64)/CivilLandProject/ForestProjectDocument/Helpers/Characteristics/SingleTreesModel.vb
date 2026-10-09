Imports System.Xml.Serialization

''' <summary>
''' Сведения о единичных деревьях (SingleTreesType)
''' </summary>
Public Class SingleTreesModel

    <XmlElement("compositionSingleTrees")>
    Public Property CompositionSingleTrees As String

    <XmlElement("averageAge")>
    Public Property AverageAge As ValueUnitMeasurementModel

    <XmlElement("treeSpecies")>
    Public Property TreeSpecies As SpeciesModel

    <XmlElement("averageStockWood")>
    Public Property AverageStockWood As ValueUnitMeasurementModel

End Class
