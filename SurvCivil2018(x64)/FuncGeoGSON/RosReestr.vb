Imports System.ComponentModel
Imports Newtonsoft.Json

<TypeConverter(GetType(ExpandableObjectConverter))>
<JsonObject>
Public Class RosReestr
    Public Property total As Integer
    Public Property features As Feature()
    Public Property total_relation As String
End Class

<TypeConverter(GetType(ExpandableObjectConverter))>
<JsonObject>
Public Class Feature
    Public Property attrs As Attrs
    Public Property sort As Long
    Public Property center As Center
    Public Property type As Integer
    Public Property extent As Extent
End Class
<TypeConverter(GetType(ExpandableObjectConverter))>
<JsonObject>
Public Class Attrs
    Public Property kvartal_cn As String 'квартал
    Public Property cn As String 'кадастровый номер
    'Public Property area_value As String 'площадь
    Public Property address As String 'адрес
    Public Property category_type As String 'категория земель
    Public Property util_code As String 'разрешенное использование по коду
    Public Property util_by_doc As String 'разрешенное использование по документу
    'Public Property cad_cost As String 'кадастровая стоимость
    Public Property statecd As String 'статус участка
    Public Property id As String 'номер частка (подавление нулей)
    'Public Property sa_status As Sa_Status
End Class

<TypeConverter(GetType(ExpandableObjectConverter))>
<JsonObject>
Public Class Sa_Status
    Public Property code As Integer
    Public Property text As String
End Class

<TypeConverter(GetType(ExpandableObjectConverter))>
<JsonObject>
Public Class Center
    Public Property y As Single
    Public Property x As Single
End Class

<TypeConverter(GetType(ExpandableObjectConverter))>
<JsonObject>
Public Class Extent
    Public Property ymax As Single
    Public Property ymin As Single
    Public Property xmax As Single
    Public Property xmin As Single
End Class

