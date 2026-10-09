Imports System.Xml.Serialization

''' <summary>
''' Сведения о лесных культурах, школах (ForestCropsType)
''' Используется и для forestCrops, и для schoolForest.
''' </summary>
Public Class ForestCropsModel

    <XmlElement("year")>
    Public Property Year As String

    <XmlElement("cultivationMethod")>
    Public Property CultivationMethod As String

    <XmlElement("creationMethod")>
    Public Property CreationMethod As String

    <XmlElement("distanceBetween")>
    Public Property DistanceBetween As ValueUnitMeasurementModel

    <XmlElement("rowSpacing")>
    Public Property RowSpacing As ValueUnitMeasurementModel

    <XmlElement("quantity")>
    Public Property Quantity As ValueUnitMeasurementModel

    <XmlElement("state")>
    Public Property State As String

    <XmlElement("causesDeath")>
    Public Property CausesDeath As String

    <XmlElement("otherInfo")>
    Public Property OtherInfo As String

End Class
