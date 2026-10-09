Imports System.Xml.Serialization

''' <summary>
''' Сведения о мероприятии (EventInfoType)
''' </summary>
Public Class EventInfoModel

    <XmlElement("eventName")>
    Public Property EventName As KindEModel

    <XmlElement("projectEventVolume")>
    Public Property ProjectEventVolume As ValueUnitMeasurementModel

    <XmlElement("rtk")>
    Public Property Rtk As String

    <XmlElement("addInfo")>
    Public Property AddInfo As String

End Class
