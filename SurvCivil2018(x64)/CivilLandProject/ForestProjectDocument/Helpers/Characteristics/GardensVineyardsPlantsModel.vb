Imports System.Xml.Serialization

''' <summary>
''' Сведения о садах, виноградниках (GardensVineyardsPlantsType)
''' </summary>
Public Class GardensVineyardsPlantsModel

    <XmlElement("gardenType")>
    Public Property GardenType As KindEModel

    <XmlElement("typeIrrigationBogarny")>
    Public Property TypeIrrigationBogarny As KindEModel

    <XmlElement("year")>
    Public Property Year As String

    <XmlElement("quantityTrees")>
    Public Property QuantityTrees As ValueUnitMeasurementModel

    <XmlElement("yield")>
    Public Property Yield As ValueUnitMeasurementModel

    <XmlElement("treeSpecies")>
    Public Property TreeSpecies As SpeciesModel

    <XmlElement("shrubbySpecies")>
    Public Property ShrubbySpecies As SpeciesModel

End Class
