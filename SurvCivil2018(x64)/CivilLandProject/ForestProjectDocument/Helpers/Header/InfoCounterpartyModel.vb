Imports System.Xml.Serialization

''' <summary>
''' Сведения о заказчике / исполнителе (InfoCounterpartyType).
''' Choice из трёх вариантов: юрлицо / учреждение / ИП.
''' Заполняется только один из трёх.
''' </summary>
Public Class InfoCounterpartyModel

    <XmlElement("infoOrganization")>
    Public Property InfoOrganization As InfoOrganizationModel

    <XmlElement("institute")>
    Public Property Institute As InstituteModel

    <XmlElement("individualEntrepreneur")>
    Public Property IndividualEntrepreneur As IndividualEntrepreneurModel

    <XmlElement("telephone")>
    Public Property Telephone As String

    <XmlElement("email")>
    Public Property Email As String

End Class
