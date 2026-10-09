Public Class RosReestr_Parcels
    Public Class Rootobject
        Public Property feature As Feature
    End Class

    Public Class Feature
        Public Property extent As Extent
        Public Property attrs As Attrs
        Public Property center As Center
        Public Property sort As Integer
        Public Property type As Integer
        Public Property stat As Stat

    End Class

    Public Class Extent
        Public Property ymin As Single
        Public Property xmin As Single
        Public Property xmax As Single
        Public Property ymax As Single
    End Class

    Public Class Attrs
        Public Property rifr_dep_info As Object
        Public Property statecd As String
        Public Property util_code As String
        Public Property area_type As String
        Public Property parcel_type As String
        Public Property area_value As Single
        Public Property sale_doc_date As Object
        Public Property rifr_dep As Object
        Public Property children As Object
        Public Property id As String
        Public Property rifr_cnt As Object
        Public Property sale_dep As Object
        Public Property kvartal_cn As String
        Public Property cad_cost As String
        Public Property cc_date_approval As String
        Public Property sale As Object
        Public Property sale_doc_num As Object
        Public Property area_unit As String
        Public Property sale_dep_uo As Object
        Public Property fp As Object
        Public Property rifr As Object
        Public Property sale_cnt As Object
        Public Property date_cost As String
        Public Property cad_unit As String
        Public Property kvartal As String
        Public Property cc_date_entering As String
        Public Property cn As String
        Public Property application_date As Object
        Public Property sale_date As Object
        Public Property sale_price As Object
        Public Property address As String
        Public Property category_type As String
        Public Property util_by_doc As String
        Public Property sale_doc_type As Object
    End Class

    Public Class Center
        Public Property y As Single
        Public Property x As Single
    End Class

    Public Class Stat
    End Class
End Class
