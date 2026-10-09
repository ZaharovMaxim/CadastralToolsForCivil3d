Imports System.Xml.Serialization

''' <summary>
''' Исходные данные (InitialDataType)
''' </summary>
Public Class InitialDataModel

    <XmlElement("documents")>
    Public Property Documents As String

    <XmlElement("infoBase")>
    Public Property InfoBase As InfoBaseModel

    <XmlElement("infoMeasuring")>
    Public Property InfoMeasuring As InfoMeasuringModel

End Class
