Imports System.Xml.Serialization

''' <summary>
''' Сведения о ярусе лесного насаждения (ForestPlantingType)
''' </summary>
Public Class ForestPlantingModel

    <XmlElement("number")>
    Public Property Number As String

    <XmlElement("mainTier")>
    Public Property MainTier As String

    <XmlElement("averageHeight")>
    Public Property AverageHeight As ValueUnitMeasurementModel

    <XmlElement("averageAge")>
    Public Property AverageAge As ValueUnitMeasurementModel

    <XmlElement("averageDiameterBarrel")>
    Public Property AverageDiameterBarrel As ValueUnitMeasurementModel

    <XmlElement("sumAreasTree")>
    Public Property SumAreasTree As ValueUnitMeasurementModel

    <XmlElement("compositionForestPlantation")>
    Public Property CompositionForestPlantation As String

    <XmlElement("completenessLayer")>
    Public Property CompletenessLayer As KindEModel

    <XmlElement("woodStockForestPlantings")>
    Public Property WoodStockForestPlantings As WoodStockForestPlantingsModel

    <XmlElement("treeSpecies")>
    Public Property TreeSpecies As SpeciesModel

    <XmlElement("shrubbySpecies")>
    Public Property ShrubbySpecies As SpeciesModel

    <XmlElement("isNaturalPlain")>
    Public Property IsNaturalPlain As String

    <XmlElement("naturalRenewal")>
    Public Property NaturalRenewal As String

End Class
