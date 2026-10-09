Imports System.Xml.Serialization
Imports System.Collections.Generic

''' <summary>
''' Сведения о средствах измерений (InfoMeasuringType)
''' </summary>
Public Class InfoMeasuringModel

    <XmlElement("Device")>
    Public Property Devices As List(Of DeviceModel)

End Class
