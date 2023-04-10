﻿Imports Newtonsoft.Json

Public Class Wrapper_OrdersCabecera
    Public Property customer_is_guest As Integer
    Public Property customer_email As String
    Public Property store_name As String
    Public Property created_at As Date
    Public Property entity_id As Integer
    Public Property increment_id As String
    Public Property coupon_code As String
    Public Property grand_total As Double
    Public Property base_shipping_amount As Double
    Public Property customer_firstname As String
    Public Property customer_lastname As String
    Public Property payment As New Wrapper_OrdersPayment
    Public Property billing_address As New Wrapper_OrdersBillingAdress
    Public Property extension_attributes As New Wrapper_OrdersExtensionAttributes
  Public Property items As New List(Of Wrapper_OrdersLineas)
  Public Property customer_id As String
End Class
