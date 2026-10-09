Imports System.Xml.Serialization

''' <summary>
''' Значение + единица измерения (ValueUnitMeasurementType)
''' </summary>
Public Class ValueUnitMeasurementModel

    <XmlElement("value")>
    Public Property Value As Decimal

    <XmlElement("unitMeasurement")>
    Public Property UnitMeasurement As KindEModel

End Class

''' <summary>
''' Площадь (sfre:AreaType)
''' </summary>
Public Class AreaModel

    <XmlElement("type")>
    Public Property Type As KindEModel

    <XmlElement("value")>
    Public Property Value As Decimal

    <XmlElement("inaccuracy")>
    Public Property Inaccuracy As Decimal?

End Class

''' <summary>
''' ФИО (cmplx:PersonNameType)
''' </summary>
Public Class PersonNameModel

    <XmlElement("last")>
    Public Property Last As String

    <XmlElement("first")>
    Public Property First As String

    <XmlElement("middle")>
    Public Property Middle As String

End Class
