Imports System.Xml.Serialization

''' <summary>
''' Сведения о подлеске (ForestUndergrowthType)
''' </summary>
Public Class ForestUndergrowthModel

    <XmlElement("composition")>
    Public Property Composition As String

    <XmlElement("averageHeight")>
    Public Property AverageHeight As ValueUnitMeasurementModel

    <XmlElement("densityUndergrowth")>
    Public Property DensityUndergrowth As KindEModel

    <XmlElement("treeSpecies")>
    Public Property TreeSpecies As SpeciesModel

    <XmlElement("shrubbySpecies")>
    Public Property ShrubbySpecies As SpeciesModel

End Class
