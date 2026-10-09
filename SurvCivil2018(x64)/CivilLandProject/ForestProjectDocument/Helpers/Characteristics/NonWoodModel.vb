Imports System.Xml.Serialization

''' <summary>
''' Сведения о недревесном сырье (NonWoodType)
''' </summary>
Public Class NonWoodModel

    <XmlElement("categoryRaw")>
    Public Property CategoryRaw As String

    <XmlElement("unit")>
    Public Property Unit As ValueUnitMeasurementModel

    <XmlElement("yield")>
    Public Property Yield As String

    <XmlElement("availability")>
    Public Property Availability As String

End Class
