Imports System.Xml.Serialization

''' <summary>
''' Сведения о подросте (TreeUndergrowthType)
''' </summary>
Public Class TreeUndergrowthModel

    <XmlElement("composition")>
    Public Property Composition As String

    <XmlElement("averageHeight")>
    Public Property AverageHeight As ValueUnitMeasurementModel

    <XmlElement("trustworthiness")>
    Public Property Trustworthiness As KindEModel

    <XmlElement("averageAge")>
    Public Property AverageAge As ValueUnitMeasurementModel

    <XmlElement("quantity")>
    Public Property Quantity As ValueUnitMeasurementModel

    <XmlElement("treeSpecies")>
    Public Property TreeSpecies As SpeciesModel

    <XmlElement("shrubbySpecies")>
    Public Property ShrubbySpecies As SpeciesModel

End Class
