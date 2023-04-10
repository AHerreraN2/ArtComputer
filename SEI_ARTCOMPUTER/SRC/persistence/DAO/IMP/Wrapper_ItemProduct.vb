Public Class Wrapper_ItemProduct
  Public Property sku As String
  Public Property name As String
  Public Property price As Double
  Public Property attribute_set_id As Integer = 159
  Public Property weight As Double = 0
  Public Property visibility = 1
  Public Property type_id As String = "simple" 'rpaya(20220329)
  Public Property extension_attributes As New Wrapper_ItemProduct_ExtensionAttributes

  Public Property status = 1 '220715 ahn

End Class

Public Class Wrapper_ItemRes
  Public Property id As String


End Class

Public Class Wrapper_Customer
  Public Property customer As New Wrapper_Client

  Public Property password As String
End Class

Public Class Wrapper_Client
  Public Property id As String
  Public Property group_id As String
  Public Property email As String
  Public Property firstname As String
  Public Property lastname As String
  Public Property website_id As String
  Public Property addresses As New List(Of Wrapper_ClientAddress)



End Class

Public Class Wrapper_ClientAddress

  Public Property id As String
  Public Property region As New Wrapper_ClientRegion
  Public Property region_id As String
  Public Property street As New List(Of String)
  Public Property company As String
  Public Property telephone As String
  Public Property fax As String
  Public Property postcode As String
  Public Property city As String
  Public Property firstname As String
  Public Property default_shipping As String
  Public Property default_billing As String





End Class

Public Class Wrapper_ClientRegion

  Public Property region_code As String
  Public Property region As String
  Public Property region_id As String


End Class

Public Class Wrapper_ClientRes
  Public Property id As String


End Class

Public Class Wrapper_BasePrice

  Public Property price As String
  Public Property store_id As String = 0
  Public Property sku As String

End Class

Public Class Wrapper_TierPrice

  Public Property price As String
  Public Property price_type As String
  Public Property website_id As String = 0
  Public Property sku As String
  Public Property customer_group As String
  Public Property quantity As String = 1


End Class