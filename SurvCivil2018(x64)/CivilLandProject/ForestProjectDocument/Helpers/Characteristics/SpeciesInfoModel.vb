Imports System.Xml.Serialization

''' <summary>
''' Сведения о древесной или кустарниковой породе (SpeciesInfoType)
''' </summary>
Public Class SpeciesInfoModel

    <XmlElement("name")>
    Public Property Name As KindEModel

    <XmlElement("shareParticipation")>
    Public Property ShareParticipation As Decimal?

    <XmlElement("age")>
    Public Property Age As ValueUnitMeasurementModel

    <XmlElement("predominantTreeSpecies")>
    Public Property PredominantTreeSpecies As String

    <XmlElement("heightTree")>
    Public Property HeightTree As ValueUnitMeasurementModel

    <XmlElement("heightShrub")>
    Public Property HeightShrub As ValueUnitMeasurementModel

    <XmlElement("averageDiameterBarrel")>
    Public Property AverageDiameterBarrel As ValueUnitMeasurementModel

    <XmlElement("ageClass")>
    Public Property AgeClass As KindEModel

    <XmlElement("ageGroupForestPlantation")>
    Public Property AgeGroupForestPlantation As KindEModel

    <XmlElement("marketabilityClass")>
    Public Property MarketabilityClass As KindEModel

    <XmlElement("stockConstituentWoodSpecies")>
    Public Property StockConstituentWoodSpecies As ValueUnitMeasurementModel

    <XmlElement("typeOrigin")>
    Public Property TypeOrigin As KindEModel

End Class
