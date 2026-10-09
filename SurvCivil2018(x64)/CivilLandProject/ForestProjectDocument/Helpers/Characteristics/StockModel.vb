Imports System.Xml.Serialization

''' <summary>
''' Запас на лесотаксационном выделе (StockType).
''' Содержит опциональные подблоки: сухостой, редины, единичные деревья,
''' неликвидная древесина, захламлённость, ликвид.
''' </summary>
Public Class StockModel

    <XmlElement("deadWood")>
    Public Property DeadWood As ValueUnitMeasurementModel

    <XmlElement("naturalRedines")>
    Public Property NaturalRedines As ValueUnitMeasurementModel

    <XmlElement("singleTrees")>
    Public Property SingleTrees As ValueUnitMeasurementModel

    <XmlElement("illiquidWood")>
    Public Property IlliquidWood As ValueUnitMeasurementModel

    <XmlElement("stockClutter")>
    Public Property StockClutter As ValueUnitMeasurementModel

    <XmlElement("liquidWood")>
    Public Property LiquidWood As ValueUnitMeasurementModel

End Class
