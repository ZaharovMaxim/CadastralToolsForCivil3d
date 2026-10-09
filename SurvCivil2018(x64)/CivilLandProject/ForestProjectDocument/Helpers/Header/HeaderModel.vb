Imports System.Xml.Serialization

''' <summary>
''' Общие сведения о документе (HeaderType)
''' </summary>
Public Class HeaderModel

    <XmlElement("infoDocument")>
    Public Property InfoDocument As InfoDocumentModel

    <XmlElement("infoBasic")>
    Public Property InfoBasic As InfoBasicModel

    <XmlElement("infoCustomer")>
    Public Property InfoCustomer As InfoCounterpartyModel

    <XmlElement("infoPerformer")>
    Public Property InfoPerformer As InfoPerformerModel

    <XmlElement("basisWork")>
    Public Property BasisWork As String

End Class