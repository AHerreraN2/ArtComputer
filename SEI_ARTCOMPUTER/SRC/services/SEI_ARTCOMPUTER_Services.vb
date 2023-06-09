Imports System.Net
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Net.Security
Imports Newtonsoft.Json
Imports System.IO
Imports Newtonsoft
Imports System.Linq
Imports System.Xml.Linq

Public Class SEI_ARTCOMPUTER_Services

#Region "Estructuras del XML"
  Public Structure Albaran
    Public Property ShipDate As String
    Public Property SalesOrderNumber As Integer
    Public Property SalesOrderDate As String
    Public Property InvoiceNumber As String
    Public Property InvoiceDate As String
    Public Property CustPO As String
    Public Property CustRef As String
    Public Property Name1_ShipTo As String
    Public Property Name2 As String
    Public Property Street As String
    Public Property ZIP As String
    Public Property City As String
    Public Property Country As String
    Public Property Name1_Buyer As String
    Public Property Name1_Seller As String
    Public Property Lines As List(Of Linea)

  End Structure

  Public Structure Linea
    Public Property ItemID_ManuFacturer As String
    Public Property ItemID_Seller As String
    Public Property ItemText As String
    Public Property ItemQty As Integer
    Public Property InvoiceLine As String
    Public Property SelesOrderLine As String
    Public Property Serials As List(Of Serial)

  End Structure

  Public Structure Serial
    Public Property Serial As String
  End Structure


#End Region

#Region "Constructor/Singleton"

  Private Shared Instance As SEI_ARTCOMPUTER_Services = Nothing

  Dim ots As New TraceSwitch("TraceSwitch", "Switch del Traceador")

  Private oCompany As SAPbobsCOM.Company = Nothing

  Public Sub New(ByRef _oCompany As SAPbobsCOM.Company)
    oCompany = _oCompany
  End Sub

  Public Shared Function GetInstance(ByRef _oCompany As SAPbobsCOM.Company) As SEI_ARTCOMPUTER_Services

    If Instance Is Nothing Then
      Instance = New SEI_ARTCOMPUTER_Services(_oCompany)
    End If

    Instance.oCompany = _oCompany

    Return Instance

  End Function

#End Region

  Public Structure Existencias
    Public Property ItemCode As String
    Public Property WhsCode As String
  End Structure


#Region "Subir Fechas de Entrega a MAGENTO (SAP a Magento)"

  Public Function UPDATEDeliveryDate_NOSYNC() As Helpers.UpdateDeliveryDateCABResponse
    Dim oResponse As Helpers.UpdateDeliveryDateCABResponse = Nothing
    Try
      oResponse = UPDATEDeliveryDateEX_NOSYNC()
    Catch ex As Exception
      Trace.WriteLineIf(ots.TraceError, ex.Message)
      Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
      oResponse = New Helpers.UpdateDeliveryDateCABResponse With {
                .ExecutionSuccess = False,
                .FailureReason = ex.Message
            }
    End Try
    Return oResponse
  End Function

  Private Function UPDATEDeliveryDateEX_NOSYNC() As Helpers.UpdateDeliveryDateCABResponse
    Dim oResponse As Helpers.UpdateDeliveryDateCABResponse = Nothing
    Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
    Dim oRecordSetUpdate As SAPbobsCOM.Recordset = Nothing
    Try
      'Buscamos los documentos pendientes de sincronizar con Magento.
      'Laura hizo una vista, en este caso para los registros de cabecera que se llama ART_MAGENTO_DELIVERYDATEGENERAL_VIEW
      Dim URL As String = vbNullString
      Dim Token As String = vbNullString
      oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

      'Buscamos la URL en la tabla de configuraci�n de la integraci�n ART_INTEGRATION
      Dim QueryURL As String = vbNullString
      Select Case oCompany.DbServerType
        Case SAPbobsCOM.BoDataServerTypes.dst_HANADB
          QueryURL = "SELECT T0.""Name"" FROM ""@ART_CONFIGURATIONS""  T0 WHERE T0.""Code"" ='URL'"
        Case Else
          QueryURL = "SELECT T0.Name FROM [@ART_CONFIGURATIONS]  T0 WHERE T0.Code ='URL'"
      End Select
      oRecordSet.DoQuery(QueryURL)
      If Not IsNothing(oRecordSet) AndAlso oRecordSet.RecordCount > 0 Then
        URL = oRecordSet.Fields.Item(0).Value.ToString.Trim
      Else
        Throw New Exception("No est� definida la URL de conexi�n con MAGENTO")
      End If
      If Not URL.EndsWith("/") Then URL = URL & "/"

      'Buscamos el BearerToken para conexi�n
      Dim QueryTOKEN As String = vbNullString
      Select Case oCompany.DbServerType
        Case SAPbobsCOM.BoDataServerTypes.dst_HANADB
          QueryTOKEN = "SELECT T0.""Name"" FROM ""@ART_CONFIGURATIONS""  T0 WHERE T0.""Code"" ='TOKEN'"
        Case Else
          QueryTOKEN = "SELECT T0.Name FROM [@ART_CONFIGURATIONS] T0 WHERE T0.Code = 'TOKEN'"
      End Select
      oRecordSet.DoQuery(QueryTOKEN)
      If Not IsNothing(oRecordSet) AndAlso oRecordSet.RecordCount > 0 Then
        Token = oRecordSet.Fields.Item(0).Value.ToString.Trim
      Else
        Throw New Exception("No est� definido el TOKEN de conexi�n con MAGENTO")
      End If

      'El DocEntry es para luego actualizar un campo en SAP. El campo WEBORDID es el ID del pedido en Magento y que usaremos en la llamada

      Dim oSQLQuery As New DNAUtils.DNASQLUtils.SQLQuery
      oSQLQuery.SelectFieldList.Add("DocEntry")
      oSQLQuery.SelectFieldList.Add("DocNum")
      oSQLQuery.SelectFieldList.Add("U_WEBORDID")
      oSQLQuery.SelectFieldList.Add("DeliveryDate")
      oSQLQuery.FromTable = "ART_MAGENTO_DELIVERYDATEGENERAL_VIEW"
      'oSQLQuery.Topper = New DNAUtils.DNASQLUtils.SQLToper(10, DNAUtils.DNASQLUtils.SQLToper.SQLToperQuantity.Value_)
      Dim Query As String = DNAUtils.DNASQLUtils.DNASQLUtils.GetInstance(oCompany).getSelectSqlSentence(oSQLQuery)

      'Dim Query As String = "SELECT ""DocEntry"", ""DocNum"", ""U_WEBORDID"", ""DeliveryDate"" FROM ""ART_MAGENTO_DELIVERYDATEGENERAL_VIEW"""
      oRecordSet.DoQuery(Query)
      oRecordSetUpdate = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
      If Not IsNothing(oRecordSet) AndAlso oRecordSet.RecordCount > 0 Then
        oRecordSet.MoveFirst()
        Dim webRequest As HttpWebRequest
        For i As Integer = 0 To oRecordSet.RecordCount - 1
          Dim DocEntry As String = oRecordSet.Fields.Item("DocEntry").Value.ToString.Trim
          Dim NumPedMag As String = oRecordSet.Fields.Item("U_WEBORDID").Value.ToString.Trim
          Dim strURL As String = URL & "V1/deliveryDate/" & NumPedMag
          Dim DocNum As String = oRecordSet.Fields.Item("DocNum").Value.ToString.Trim
          Dim DeliveryCabecera As New Wrapper_DeliveryDateCabecera
          DeliveryCabecera.deliveryDate = oRecordSet.Fields.Item("DeliveryDate").Value.ToString.Trim

          Dim ServiceCall As String = JsonConvert.SerializeObject(DeliveryCabecera)
          Trace.WriteLineIf(ots.TraceInfo, "URL de Servicio: " & strURL)
          Trace.WriteLineIf(ots.TraceInfo, "ServiceCall CABECERA Pedidos: " & ServiceCall)
          Dim data = Encoding.UTF8.GetBytes(ServiceCall)

          '    'Llamamos al SW
          ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf AcceptAllCertifications)

          ServicePointManager.Expect100Continue = True
          ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

          webRequest = HttpWebRequest.Create(strURL)
          webRequest.AllowAutoRedirect = True
          webRequest.Timeout = -1
          webRequest.Method = WebRequestMethods.Http.Post
          webRequest.ContentLength = data.Length
          webRequest.ContentType = "application/json"
          webRequest.Headers.Set("Authorization", "Bearer " & Token)
          Dim stream = webRequest.GetRequestStream()
          stream.Write(data, 0, data.Length)
          stream.Close()
          stream.Dispose()
          'Hacemos la llamada al servicio
          Try
            Dim webResponse As HttpWebResponse = webRequest.GetResponse()
            If webResponse.StatusCode = HttpStatusCode.OK Then
              Trace.WriteLineIf(ots.TraceInfo, "Fecha de Entrega del Pedido " & DocNum & " actualizado en MAGENTO")
            End If
            webResponse.Close()
          Catch ex As Exception
            RECORD_ERRORS(ex.Message, "UPDATE ORDER DELIVERY DATE->DOCUMENT")
            Trace.WriteLineIf(ots.TraceError, ex.Message)
            Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
          Finally
          End Try
          stream = Nothing
          oRecordSet.MoveNext()
        Next
        'Actualizamos las fechas de las l�neas del pedido en MAGENTO
        oSQLQuery = Nothing
        oSQLQuery = New DNAUtils.DNASQLUtils.SQLQuery
        oSQLQuery.SelectFieldList.Add("DocEntry")
        oSQLQuery.SelectFieldList.Add("DocNum")
        oSQLQuery.SelectFieldList.Add("U_WEBORDID")
        oSQLQuery.SelectFieldList.Add("DeliveryDate")
        oSQLQuery.SelectFieldList.Add("WebOrderItemId")
        oSQLQuery.SelectFieldList.Add("ItemCode")
        oSQLQuery.FromTable = "ART_MAGENTO_DELIVERYDATELINE_VIEW"
        'oSQLQuery.Topper = New DNAUtils.DNASQLUtils.SQLToper(10, DNAUtils.DNASQLUtils.SQLToper.SQLToperQuantity.Value_)
        Query = DNAUtils.DNASQLUtils.DNASQLUtils.GetInstance(oCompany).getSelectSqlSentence(oSQLQuery)
        'Query = "SELECT ""DocNum"", ""DocEntry"", ""U_WEBORDID"", ""WebOrderItemId"", ""ItemCode"", ""DeliveryDate"", ""LineNum"" FROM ""ART_MAGENTO_DELIVERYDATELINE_VIEW"""
        oRecordSet.DoQuery(Query)
        If Not IsNothing(oRecordSet) AndAlso oRecordSet.RecordCount > 0 Then
          Dim ListaDocumentos As New Dictionary(Of String, List(Of String)) 'Lista con DocEntry, WebOrderItemId que permitir� actualizar el campo U_UpdateDDMagento ='N' en la l�nea del Pedido. (El campo WebOrderItemId es como el LineNum en Magento, as� que no se repite en un documento)
          oRecordSet.MoveFirst()
          For i As Integer = 0 To oRecordSet.RecordCount - 1
            Dim DocEntry As String = oRecordSet.Fields.Item("DocEntry").Value.ToString.Trim
            Dim NumPedMag As String = oRecordSet.Fields.Item("U_WEBORDID").Value.ToString.Trim
            Dim ItemCode As String = oRecordSet.Fields.Item("ItemCode").Value.ToString.Trim
            Dim ArtMagento As String = oRecordSet.Fields.Item("WebOrderItemId").Value.ToString.Trim
            Dim DocNum As String = oRecordSet.Fields.Item("DocNum").Value.ToString.Trim
            Dim DeliveryDateLinea As String = oRecordSet.Fields.Item("DeliveryDate").Value.ToString.Trim
            Dim strURL As String = URL & "V1/deliveryDateItem/" & NumPedMag
            Dim DeliveryLinea As New Wrapper_DeliveryDateLineas
            DeliveryLinea.itemId = ArtMagento
            DeliveryLinea.deliveryDate = DeliveryDateLinea

            Dim ServiceCall As String = JsonConvert.SerializeObject(DeliveryLinea)
            Trace.WriteLineIf(ots.TraceInfo, "URL de Servicio: " & strURL)
            Trace.WriteLineIf(ots.TraceInfo, "ServiceCall LINEAS Pedidos: " & ServiceCall)
            Dim data = Encoding.UTF8.GetBytes(ServiceCall)

            '    'Llamamos al SW
            ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf AcceptAllCertifications)

            ServicePointManager.Expect100Continue = True
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            webRequest = HttpWebRequest.Create(strURL)
            webRequest.AllowAutoRedirect = True
            webRequest.Timeout = -1
            webRequest.Method = WebRequestMethods.Http.Post
            webRequest.ContentLength = data.Length
            webRequest.ContentType = "application/json"
            webRequest.Headers.Set("Authorization", "Bearer " & Token)
            Dim stream = webRequest.GetRequestStream()
            stream.Write(data, 0, data.Length)
            stream.Close()
            stream.Dispose()
            Try
              Dim webResponse As HttpWebResponse = webRequest.GetResponse()
              If webResponse.StatusCode = HttpStatusCode.OK Then
                Trace.WriteLineIf(ots.TraceInfo, "Fecha de Entrega del Pedido " & DocNum & " actualizado en MAGENTO")
                'Agregamos la l�nea que s� subi� a MAGENTO
                'Chequeamos si el docentry existe. De ser as�, agregamos la l�nea. Si no, agregamos el docentry + la l�nea
                If ListaDocumentos.Count = 0 Then
                  Dim ListaLineNum As New List(Of String)
                  ListaLineNum.Add(ArtMagento)
                  ListaDocumentos.Add(DocEntry, ListaLineNum)
                Else
                  If ListaDocumentos.ContainsKey(DocEntry) Then
                    'Si est� el DocEntry, agrego la l�nea en el arreglo
                    Dim ListofLineNum As List(Of String) = ListaDocumentos.Item(DocEntry)
                    ListofLineNum.Add(ArtMagento)
                    ListaDocumentos.Item(DocEntry) = ListofLineNum
                  Else
                    Dim ListaLineNum As New List(Of String)
                    ListaLineNum.Add(ArtMagento)
                    ListaDocumentos.Add(DocEntry, ListaLineNum)
                  End If

                End If
              End If
              webResponse.Close()
            Catch ex As Exception
              RECORD_ERRORS(ex.Message, "UPDATE ORDER DELIVERY DATE->LINES")
              Trace.WriteLineIf(ots.TraceError, ex.Message)
              Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
            Finally
            End Try
            stream = Nothing
            oRecordSet.MoveNext()
          Next
          If ListaDocumentos.Count > 0 Then
            UPDATE_ORDERLINE_STATUS(ListaDocumentos)
          End If
        Else
          Trace.WriteLineIf(ots.TraceInfo, "No hay l�neas de Pedidos pendientes de sincronizar con magento")
        End If
      Else
        Trace.WriteLineIf(ots.TraceInfo, "No hay Pedidos pendientes de sincronizar con magento")
      End If
      oResponse = New Helpers.UpdateDeliveryDateCABResponse
      oResponse.ExecutionSuccess = True
      oResponse.FailureReason = String.Empty
    Catch ex As Exception
      Trace.WriteLineIf(ots.TraceError, ex.Message)
      Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
      oResponse = New Helpers.UpdateDeliveryDateCABResponse
      oResponse.ExecutionSuccess = False
      oResponse.FailureReason = ex.Message
    Finally
      If Not IsNothing(oRecordSet) Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
      If Not IsNothing(oRecordSetUpdate) Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSetUpdate)
    End Try
    Return oResponse
  End Function

#End Region

#Region "Subir stocks (SAP a Magento)"

  Public Function UPDATEStocks_NOSYNC() As Helpers.UpdateStocksResponse
    Dim oResponse As Helpers.UpdateStocksResponse = Nothing
    Try
      oResponse = UPDATEStocksEX_NOSYNC()
    Catch ex As Exception
      Trace.WriteLineIf(ots.TraceError, ex.Message)
      Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
      oResponse = New Helpers.UpdateStocksResponse With {
                .ExecutionSuccess = False,
                .FailureReason = ex.Message
            }
    End Try
    Return oResponse
  End Function

  Private Function UPDATEStocksEX_NOSYNC() As Helpers.UpdateStocksResponse
    Dim oResponse As Helpers.UpdateStocksResponse
    Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
    Dim oRecordSetUpdate As SAPbobsCOM.Recordset = Nothing
    Try
      Dim URL As String = vbNullString
      Dim Token As String = vbNullString
      oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
      'Buscamos la URL en la tabla de configuraci�n de la integraci�n ART_INTEGRATION
      Dim QueryURL As String = vbNullString
      Select Case oCompany.DbServerType
        Case SAPbobsCOM.BoDataServerTypes.dst_HANADB
          QueryURL = "SELECT T0.""Name"" FROM ""@ART_CONFIGURATIONS""  T0 WHERE T0.""Code"" ='URL'"
        Case Else
          QueryURL = "SELECT T0.Name FROM [@ART_CONFIGURATIONS]  T0 WHERE T0.Code ='URL'"
      End Select
      oRecordSet.DoQuery(QueryURL)
      If Not IsNothing(oRecordSet) AndAlso oRecordSet.RecordCount > 0 Then
        URL = oRecordSet.Fields.Item(0).Value.ToString.Trim
      Else
        Throw New Exception("No est� definida la URL de conexi�n con MAGENTO")
      End If
      If Not URL.EndsWith("/") Then URL = URL & "/"

      'Buscamos el BearerToken para conexi�n
      Dim QueryTOKEN As String = vbNullString
      Select Case oCompany.DbServerType
        Case SAPbobsCOM.BoDataServerTypes.dst_HANADB
          QueryTOKEN = "SELECT T0.""Name"" FROM ""@ART_CONFIGURATIONS""  T0 WHERE T0.""Code"" ='TOKEN'"
        Case Else
          QueryTOKEN = "SELECT T0.Name FROM [@ART_CONFIGURATIONS] T0 WHERE T0.Code = 'TOKEN'"
      End Select
      oRecordSet.DoQuery(QueryTOKEN)
      If Not IsNothing(oRecordSet) AndAlso oRecordSet.RecordCount > 0 Then
        Token = oRecordSet.Fields.Item(0).Value.ToString.Trim
      Else
        Throw New Exception("No est� definido el TOKEN de conexi�n con MAGENTO")
      End If
      '
      Dim oSQLQuery As New DNAUtils.DNASQLUtils.SQLQuery
      oSQLQuery.SelectFieldList.Add("SuppCatNum")
      oSQLQuery.SelectFieldList.Add("U_WebID")
      oSQLQuery.SelectFieldList.Add("TotalStock")
      oSQLQuery.SelectFieldList.Add("ItemCode")
      oSQLQuery.FromTable = "ART_MAGENTO_STOCKVIEW"
      'oSQLQuery.Topper = New DNAUtils.DNASQLUtils.SQLToper(10, DNAUtils.DNASQLUtils.SQLToper.SQLToperQuantity.Value_)
      Dim Query As String = DNAUtils.DNASQLUtils.DNASQLUtils.GetInstance(oCompany).getSelectSqlSentence(oSQLQuery)

      Dim ListItemCode As New Dictionary(Of String, Double)
      oRecordSet.DoQuery(Query)
      If Not IsNothing(oRecordSet) AndAlso oRecordSet.RecordCount > 0 Then
        oRecordSet.MoveFirst()
        Dim List As New List(Of Wrapper_StockItems)
        For i As Integer = 0 To oRecordSet.RecordCount - 1
          Dim Items As New Wrapper_StockItems
          Dim productSku As String = oRecordSet.Fields.Item(3).Value.ToString.Trim
          Dim itemId As String = oRecordSet.Fields.Item(1).Value.ToString.Trim
          Dim qty As Double = oRecordSet.Fields.Item(2).Value
          Dim ItemCode As String = oRecordSet.Fields.Item(0).Value
          ListItemCode.Add(productSku, qty)
          Items.productSku = productSku
          Dim Item As New Wrapper_StockItem
          Item.itemId = itemId
          Item.qty = qty
          Items.stockItem = Item
          List.Add(Items)
          oRecordSet.MoveNext()
        Next
        '    'Llamamos al SW
        ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf AcceptAllCertifications)

        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

        Dim ServiceCall As String = JsonConvert.SerializeObject(List)
        Trace.WriteLineIf(ots.TraceInfo, "ServiceCall: " & ServiceCall)
        Dim data = Encoding.UTF8.GetBytes(ServiceCall)
        Dim strURL As String = URL & "async/bulk/V1/products/byProductSku/stockItems/byItemId"
        Trace.WriteLineIf(ots.TraceInfo, "EndPoint: " & strURL)
        Dim webRequest As HttpWebRequest
        webRequest = HttpWebRequest.Create(strURL)
        webRequest.AllowAutoRedirect = True
        webRequest.Timeout = -1
        webRequest.Method = WebRequestMethods.Http.Put
        webRequest.ContentLength = data.Length
        webRequest.ContentType = "application/json"
        webRequest.Headers.Set("Authorization", "Bearer " & Token)
        Dim stream = webRequest.GetRequestStream()
        stream.Write(data, 0, data.Length)
        stream.Close()
        stream.Dispose()

        Try
          Dim webResponse As HttpWebResponse = webRequest.GetResponse()
          If webResponse.StatusCode = HttpStatusCode.Accepted Then
            Trace.WriteLineIf(ots.TraceInfo, "Stocks actualizados en MAGENTO")
          End If
          webResponse.Close()
          'Ahora actualizamos los registros en SAP para que no vuelvan a subir
          If Not UPDATE_ITEMSTATUS(ListItemCode) Then Throw New Exception("Error al actualizar el estatus de los art�culos en SAP")
        Catch ex As Exception
          RECORD_ERRORS(ex.Message, "UPDATE STOCKS")
          Trace.WriteLineIf(ots.TraceError, ex.Message)
          Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
        Finally
        End Try
        stream = Nothing
      Else
        Trace.WriteLineIf(ots.TraceInfo, "No hay art�culos pendientes de sincronizar su stock con magento")
      End If
      oResponse = New Helpers.UpdateStocksResponse
      oResponse.ExecutionSuccess = True
      oResponse.FailureReason = String.Empty
    Catch ex As Exception
      Trace.WriteLineIf(ots.TraceError, ex.Message)
      Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
      RECORD_ERRORS(ex.Message, "UPDATE STOCKS")
      oResponse = New Helpers.UpdateStocksResponse
      oResponse.ExecutionSuccess = False
      oResponse.FailureReason = ex.Message
    End Try
    Return oResponse
  End Function

#End Region

#Region "Descargar Pedidos (Magento a SAP)"

  Public Function DOWNLOAD_SalesOrderNOSYNC() As Helpers.DownloadSalesOrdersResponse
    Dim oResponse As Helpers.DownloadSalesOrdersResponse = Nothing
    Try
      oResponse = DOWNLOAD_SalesOrderEX_NOSYNC()
    Catch ex As Exception
      Trace.WriteLineIf(ots.TraceError, ex.Message)
      Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
      oResponse = New Helpers.DownloadSalesOrdersResponse With {
            .ExecutionSuccess = False,
            .FailureReason = ex.Message
            }
    End Try
    Return oResponse
  End Function

  Private Function DOWNLOAD_SalesOrderEX_NOSYNC() As Helpers.DownloadSalesOrdersResponse
    Dim oResponse As Helpers.DownloadSalesOrdersResponse = Nothing
    Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
    Try
      Dim URL As String = vbNullString
      Dim Token As String = vbNullString
      oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
      'Buscamos la URL en la tabla de configuraci�n de la integraci�n ART_INTEGRATION
      Dim QueryURL As String = vbNullString
      Select Case oCompany.DbServerType
        Case SAPbobsCOM.BoDataServerTypes.dst_HANADB
          QueryURL = "SELECT T0.""Name"" FROM ""@ART_CONFIGURATIONS""  T0 WHERE T0.""Code"" ='URL'"
        Case Else
          QueryURL = "SELECT T0.Name FROM [@ART_CONFIGURATIONS]  T0 WHERE T0.Code ='URL'"
      End Select
      oRecordSet.DoQuery(QueryURL)
      If Not IsNothing(oRecordSet) AndAlso oRecordSet.RecordCount > 0 Then
        URL = oRecordSet.Fields.Item(0).Value.ToString.Trim
      Else
        Throw New Exception("The connection URL with MAGENTO is not defined")
      End If
      If Not URL.EndsWith("/") Then URL = URL & "/"

      'Buscamos el BearerToken para conexi�n
      Dim QueryTOKEN As String = vbNullString
      Select Case oCompany.DbServerType
        Case SAPbobsCOM.BoDataServerTypes.dst_HANADB
          QueryTOKEN = "SELECT T0.""Name"" FROM ""@ART_CONFIGURATIONS""  T0 WHERE T0.""Code"" ='TOKEN'"
        Case Else
          QueryTOKEN = "SELECT T0.Name FROM [@ART_CONFIGURATIONS] T0 WHERE T0.Code = 'TOKEN'"
      End Select
      oRecordSet.DoQuery(QueryTOKEN)
      If Not IsNothing(oRecordSet) AndAlso oRecordSet.RecordCount > 0 Then
        Token = oRecordSet.Fields.Item(0).Value.ToString.Trim
      Else
        Throw New Exception("The Token for connect to MAGENTO is not defined")
      End If

      'Hay un paso inicial que es llamar a un EndPoint que nos devolver� un listado con todos los pedidos pendientes de bajar desde Magento. (Sergio lo indica en correo del 10/02/2022 a las 11:01h)
      'Esto nos va a devolver un listado de n�mero de pedidos. Luego con ese listado lo que tenemos que hacer es llamar tantas veces como pedidos existan al endpoint https://mcstaging-cern.artcomputer.ch/rest/V1/orders/ indicando el n�mero de pedido para sacar los detalles.

      'Paso 1: Llamada al endpoint para sacar listado de Pedidos pendientes de bajar

      '    'Llamamos al SW
      ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf AcceptAllCertifications)
      ServicePointManager.Expect100Continue = True
      ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

      Trace.WriteLineIf(ots.TraceInfo, "--------------------------PROCESS BEGINNING---------------------------------")
      Dim webRequest As HttpWebRequest
      Dim strURL As String = URL & "V1/sapSync/getUnsyncs"
      webRequest = HttpWebRequest.Create(strURL)
      webRequest.AllowAutoRedirect = True
      webRequest.Timeout = -1
      webRequest.Method = WebRequestMethods.Http.Get
      webRequest.ContentType = "application/json"
      webRequest.Headers.Set("Authorization", "Bearer " & Token)
      Dim webResponse As HttpWebResponse = webRequest.GetResponse()
      Dim receiveStreamListadoPedidos As Stream = webResponse.GetResponseStream
      Dim readStreamListadoPedidos As StreamReader = New StreamReader(receiveStreamListadoPedidos, Encoding.UTF8)
      Dim respuestaListadoPedidos As String = readStreamListadoPedidos.ReadToEnd
      Dim ListadoDocs As New List(Of String)
      ListadoDocs = JsonConvert.DeserializeObject(Of List(Of String))(respuestaListadoPedidos)
      If ListadoDocs.Count > 0 Then
        Trace.WriteLineIf(ots.TraceInfo, "There are " & ListadoDocs.Count & " sales orders pending to sincronize with SAP")
        strURL = URL & "V1/orders/"
        'Paso 2: Recorrer el listado de documentos e ir haciendo una a una las llamadas.
        For Each Documento In ListadoDocs
          Trace.WriteLineIf(ots.TraceInfo, "Downloading information for Sales Order: " & Documento)
          webRequest = HttpWebRequest.Create(strURL & Documento)
          webRequest.AllowAutoRedirect = True
          webRequest.Timeout = -1
          webRequest.Method = WebRequestMethods.Http.Get
          webRequest.ContentType = "application/json"
          webRequest.Headers.Set("Authorization", "Bearer " & Token)
          Dim oOrder As SAPbobsCOM.Documents = Nothing
          Dim oRs As SAPbobsCOM.Recordset = Nothing
          If Not oCompany.InTransaction Then
            oCompany.StartTransaction()
          End If
          Try
            webResponse = webRequest.GetResponse()
            If webResponse.StatusCode = HttpStatusCode.OK Then
              Dim receiveStream As Stream = webResponse.GetResponseStream
              Dim readStream As StreamReader = New StreamReader(receiveStream, Encoding.UTF8)
              Dim respuesta As String = readStream.ReadToEnd
              Dim sCorreo As String = String.Empty
              'Dim Resp = JsonConvert.DeserializeObject(respuesta)
              'Dim Billing_street As String = Resp("billing_address")("street").ToString.Replace("[", "").Replace("]", "")
              'Dim Shipping_street As String = Resp("extension_attributes")("shipping_assignments")(0)("shipping")("address")("street").ToString.Replace("[", "").Replace("]", "")
              Dim oPedido As New Wrapper_OrdersCabecera
              oPedido = JsonConvert.DeserializeObject(Of Wrapper_OrdersCabecera)(respuesta)
              Trace.WriteLineIf(ots.TraceInfo, "MAGENTO Answer: " & respuesta)
              oOrder = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
              'Modificaciones enviadas por Laura el 21/02/2022 a las 09:00h. 
              'Si la etiqueta store_name empieza en CERN, entonces usar el cliente C00007. Sino, dejar las reglas que ya estaban
              Dim sStoreName As String = oPedido.store_name
              If sStoreName.StartsWith("CERN") Then
                oOrder.CardCode = "C00007"
              Else
                'Correo Laura 08/02/2022 10:16h. En ese correo est�n todas las especificaciones de los campos a leer del JSON y su equivalencia con SAP.
                If oPedido.customer_is_guest = 1 Then
                  'Si la etiqueta "customer_is_guest" tiene el valor 1 entonces usar el CardCode 'C63062'. 
                  'JASF (22/06/2022) Se cambia el c�digo por C63032
                  oOrder.CardCode = "C63032"
                Else
                  ' Sino entonces buscar el cliente en SAP que tenga asociada la persona de contacto con la direcci�n de correo de "customer_email", que est� activa, que sea un cliente (OCRD.CardType='C'), y que el cliente que tenga asociado est� activo y que el grupo del cliente sea:
                  '�	Si la etiqueta "store_name" empieza por B2C, el grupo de cliente debe ser 103, 113 
                  '�	Si la etiqueta "store_name" empieza por B2B, CERN, el grupo de cliente no debe ser 103, 113
                  '�	Sino se obtiene nada, entonces el correo en cualquier persona de contacto y si no hay entonces buscar el email en el cliente 
                  ' SI NINGUN CLIENTE EXISTE, crearlo - -> con las reglas que se especifican en el cuadro MAGENTO->SAP: Customers

                  'Campo SAP  /  Campo Magento
                  'DocDate	Created-at
                  'DocDueDate   Created-at
                  'NumAtCard	Po_number
                  'U_Coupon	Coupon_code
                  'U_WEBORDID	Entity_id
                  'U_NumAtCard	Increment_id
                  'U_PAYMETH	Payment-method
                  'Billing address	Billing address details
                  'Shipping address	Shipping address details
                  'U_DEP (Y/N)	DEP field in the address (false/true)
                  'U_WEBORDR	'Y'
                  'U_DiffAmount_SAPMag	Despu�s de crear la SO, verificar si el campo de total enviado en la orden (base_total_invoiced) es igual que el de SAP, y si no lo es colocar este campo en Y

                  'LINEAS	 

                  'Correo de Laura el 15/02/2022 12:39:
                  '   o	Ignorar el art�culo configurable (product_type = �configurable�)
                  '   o	Del art�culo simple (hijo) tomar: item_id, qty_ordered, sku
                  '   o   Dentro del segmento "parent_item" tomar la informaci�n restante:precio, linetotal, etc. Si el segmento "parent_item" no existe, debemos buscar la mismo info pero en el segmento superior.

                  'U_WebOrderItemId	Item_id
                  'ItemCode	Sku
                  'UnitPrice	base_original_price
                  'LineTotal	row_total
                  'Quantity	Qty_ordered
                  'DiscountPercent	Discount_amount
                  'ExpenseCode	Si el campo "base_shipping_amount" es diferente de cero, hacer :
                  '�	A�adir un porte con el codigo "DeliveryCost" C�digo 1
                  '�	Con "Net amount" = "base_shipping_amount"
                  'OcrCode	�	Si el "store_name" empieza con B2C --> 'E-com002'
                  '�	Sino --> 'HQ003'
                  'DocDueDate	Tomar la fecha que se estima de entrega en Magento (espercificar)

                  'Warehouse 	�	 Si el campo OSCN.U_OnHand del art�culo tiene cantidad superior o igual a la cantidad pedida en la l�nea del pedido, entonces colocar el pedido en el almac�n 11.  
                  '�	Si hay stock (OnHand - IsCommited) en el almac�n 01, colocarlo en ese almac�n, y sino, colocarlo en cualquier almac�n de tienda (OWSH.U_WebStock='Y') donde haya stock, por orden de cantidad y que pueda servir la cantidad de la l�nea del pedido
                  '�	Si no se cumple ninguna de las condiciones anteriores, entonces dejar de todas todas el almac�n 01
                  'OJO: �ser� mejor hacer una vista, funci�n o procedimiento almacenado, en el que env�es el c�digo del art�culo y se devuelva el almac�n?



                  sCorreo = oPedido.customer_email
                  If String.IsNullOrEmpty(sCorreo) Then Throw New Exception("Customer doesn't have its email address")
                  If String.IsNullOrEmpty(sStoreName) Then Throw New Exception("Store name is empty")
                  Dim sCondicion As String = String.Empty
                  If sStoreName.StartsWith("B2C") Then sCondicion = " AND T0.""GroupCode"" IN (103,113)"
                  If sStoreName.StartsWith("B2B") Or sStoreName.StartsWith("CERN") Then sCondicion = " AND T0.""GroupCode"" NOT IN (103,113)"
                  Dim sQuery As String = "SELECT T0.""CardCode"" 
                                                            FROM
                                                            OCRD T0
                                                            INNER JOIN OCPR T1 ON T0.""CardCode"" = T1.""CardCode""
                                                            WHERE T1.""E_MailL""='" & sCorreo & "' AND T1.""Active"" = 'Y' AND T0.""validFor"" = 'Y' AND T0.""CardType"" = 'C' "
                  oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                  oRs.DoQuery(sQuery & sCondicion)
                  If Not IsNothing(oRs) AndAlso oRs.RecordCount > 0 Then
                    If oRs.RecordCount > 1 Then Throw New Exception("There are more than one customer in SAP with the same email: " & sCorreo)
                    oOrder.CardCode = oRs.Fields.Item(0).Value.ToString.Trim
                  Else
                    oRs.DoQuery(sQuery)
                    If oRs.RecordCount > 0 Then
                      If oRs.RecordCount > 1 Then Throw New Exception("There are more than one customer in SAP with the same email: " & sCorreo)
                      oOrder.CardCode = oRs.Fields.Item(0).Value.ToString.Trim
                    Else
                      sQuery = "SELECT T0.""CardCode""
                                                FROM 
                                                OCRD T0
                                                WHERE T0.""E_Mail""='" & sCorreo & "' AND T0.""validFor""='Y' AND T0.""CardType""='C'"
                      oRs.DoQuery(sQuery)
                      If oRs.RecordCount > 0 Then
                        If oRs.RecordCount > 1 Then
                          Throw New Exception("There are more than one customer in SAP with the same email: " & sCorreo)
                        Else
                          oOrder.CardCode = oRs.Fields.Item(0).Value.ToString.Trim
                        End If
                      Else
                        Dim NuevoCliente As String = ADD_NEW_CUSTOMER(oPedido)
                        If NuevoCliente = "-1" Then Throw New Exception("Error adding the new customer...check log for details")
                        oOrder.CardCode = NuevoCliente
                      End If
                    End If
                  End If

                End If


              End If
              If String.IsNullOrEmpty(oOrder.CardCode) Then Throw New Exception("No customer found with email: " & sCorreo)
              oOrder.DocDate = oPedido.created_at
              oOrder.NumAtCard = oPedido.extension_attributes.pgw.po_number

              Trace.WriteLineIf(ots.TraceInfo, "Sales Order's head fields added")

              'Campos de Usuario
              If Not IsNothing(oPedido.coupon_code) AndAlso Not String.IsNullOrEmpty(oPedido.coupon_code) Then
                oOrder.UserFields.Fields.Item("U_Coupon").Value = oPedido.coupon_code
              End If
              oOrder.UserFields.Fields.Item("U_WEBORDID").Value = CStr(oPedido.entity_id)
              'JASF: Modificado el 18/04/2022 por correo de Jonathan el 06/04/2022 11:02. Si el "store_name" NO empieza con Cern, debe ser Y. Caso contrario tendr� valor N
              If Not sStoreName.StartsWith("CERN") Then
                oOrder.UserFields.Fields.Item("U_WEBORDR").Value = "Y"
              Else
                oOrder.UserFields.Fields.Item("U_WEBORDR").Value = "N"
              End If

              oOrder.UserFields.Fields.Item("U_NumAtCard").Value = oPedido.increment_id
              oOrder.UserFields.Fields.Item("U_PAYMETH").Value = oPedido.payment.method
              oOrder.UserFields.Fields.Item("U_MagentoDocTotal").Value = oPedido.grand_total
              Select Case oPedido.extension_attributes.dep.dep
                Case True
                  oOrder.UserFields.Fields.Item("U_DEP").Value = "Y"
                Case False
                  oOrder.UserFields.Fields.Item("U_DEP").Value = "N"
              End Select
              Trace.WriteLineIf(ots.TraceInfo, "SAP User's fields added")

              'Direcci�n de Facturaci�n
              Dim DirFact As String = String.Empty
              For Each part In oPedido.billing_address.street
                DirFact += part & " "
              Next
              If DirFact.Length > 100 Then DirFact = DirFact.Substring(0, 99)

              Dim squery2 = "select * from crd1 Where ""CardCode"" = '" & oOrder.CardCode & "' and ""U_WBCUSTADDID"" = '" & oPedido.billing_address.customer_address_id & "' AND ""AdresType"" = 'B'"
              oRs.DoQuery(squery2)
              If Not oRs.EoF Then
                oOrder.Address = oRs.Fields.Item("Address").Value
              Else

              End If

              squery2 = "select * from crd1 Where ""CardCode"" = '" & oOrder.CardCode & "' and ""U_WBCUSTADDID"" = '" & oPedido.extension_attributes.shipping_assignments(0).shipping.address.customer_address_id & "' AND ""AdresType"" = 'S'"
              oRs.DoQuery(squery2)
              If Not oRs.EoF Then
                oOrder.Address2 = oRs.Fields.Item("Address").Value
              Else

              End If
              '220715
              ''oOrder.Address = "Billing"
              ''oOrder.Address2 = "Shipping"


              ''oOrder.AddressExtension.BillToStreet = DirFact
              ''oOrder.AddressExtension.BillToCountry = oPedido.billing_address.country_id
              ''oOrder.AddressExtension.BillToCity = oPedido.billing_address.city
              ''oOrder.AddressExtension.BillToZipCode = oPedido.billing_address.postcode
              ''oOrder.AddressExtension.BillToState = oPedido.billing_address.region_code
              ''oOrder.AddressExtension.BillToBlock = oPedido.billing_address.telephone
              ''oOrder.AddressExtension.BillToAddress2 = oPedido.billing_address.company
              ''oOrder.AddressExtension.BillToAddress3 = oPedido.billing_address.firstname & " " & oPedido.billing_address.lastname

              ''Trace.WriteLineIf(ots.TraceInfo, "Billing Address Added")

              '''Direcci�n de Entrega
              ''Dim DirEnv As String = String.Empty
              ''For Each part In oPedido.extension_attributes.shipping_assignments(0).shipping.address.street
              ''  DirEnv += part & " "
              ''Next
              ''If DirEnv.Length > 100 Then DirEnv = DirEnv.Substring(0, 99)



              ''oOrder.AddressExtension.ShipToStreet = DirEnv
              ''oOrder.AddressExtension.ShipToCountry = oPedido.extension_attributes.shipping_assignments(0).shipping.address.country_id
              ''oOrder.AddressExtension.ShipToCity = oPedido.extension_attributes.shipping_assignments(0).shipping.address.city
              ''oOrder.AddressExtension.ShipToZipCode = oPedido.extension_attributes.shipping_assignments(0).shipping.address.postcode
              ''oOrder.AddressExtension.ShipToState = oPedido.extension_attributes.shipping_assignments(0).shipping.address.region_code
              ''oOrder.AddressExtension.ShipToBlock = oPedido.extension_attributes.shipping_assignments(0).shipping.address.telephone
              ''oOrder.AddressExtension.ShipToAddress2 = oPedido.extension_attributes.shipping_assignments(0).shipping.address.company
              ''oOrder.AddressExtension.ShipToAddress3 = oPedido.extension_attributes.shipping_assignments(0).shipping.address.firstname & " " & oPedido.extension_attributes.shipping_assignments(0).shipping.address.lastname

              ''Trace.WriteLineIf(ots.TraceInfo, "Shipping Address Added")

              'Gastos de Env�o
              If Not IsNothing(oPedido.base_shipping_amount) AndAlso oPedido.base_shipping_amount > 0 Then
                oOrder.Expenses.ExpenseCode = 1
                oOrder.Expenses.LineTotal = oPedido.base_shipping_amount
                oOrder.Expenses.Add()
              End If

              Trace.WriteLineIf(ots.TraceInfo, "Additional Expenses Added (Delivery Costs)")

              ' Persona de contacto, si se tiene una persona de contacto asociada al correo con el que identific� al cliente, seleccionar esa persona de contacto como asociada al documento. 
              If Not String.IsNullOrEmpty(sCorreo) Then
                Dim CntcCode As Integer = RETURN_CONTACTPERSON(oOrder.CardCode, sCorreo)
                If CntcCode <> -1 Then oOrder.ContactPersonCode = CntcCode
              End If

              Trace.WriteLineIf(ots.TraceInfo, "Contact Person Added")

              Trace.WriteLineIf(ots.TraceInfo, "Adding Lines...")
              Dim DeliveryDate As Date = Now.Date
              For Each linea In oPedido.items
                If linea.parent_item.product_type = "configurable" Then Continue For
                If oOrder.Lines.ItemCode <> "" Then oOrder.Lines.Add()
                If Not IsNothing(linea.parent_item) AndAlso Not String.IsNullOrEmpty(linea.parent_item.sku) Then
                  oOrder.Lines.ItemCode = linea.parent_item.sku
                Else
                  oOrder.Lines.ItemCode = linea.sku
                End If
                If Not IsNothing(linea.parent_item) AndAlso linea.parent_item.item_id <> 0 Then
                  oOrder.Lines.UserFields.Fields.Item("U_WebOrderItemId").Value = CStr(linea.parent_item.item_id)
                Else
                  oOrder.Lines.UserFields.Fields.Item("U_WebOrderItemId").Value = CStr(linea.item_id)
                End If
                If Not IsNothing(linea.parent_item) AndAlso linea.parent_item.base_original_price <> 0 Then
                  oOrder.Lines.UnitPrice = linea.parent_item.base_original_price
                Else
                  oOrder.Lines.UnitPrice = linea.base_original_price
                End If
                If Not IsNothing(linea.parent_item) AndAlso linea.parent_item.row_total <> 0 Then
                  oOrder.Lines.LineTotal = linea.parent_item.row_total + linea.discount_tax_compensation_amount
                Else
                  oOrder.Lines.LineTotal = linea.row_total + linea.discount_tax_compensation_amount
                End If
                If Not IsNothing(linea.parent_item) AndAlso linea.parent_item.qty_ordered <> 0 Then
                  oOrder.Lines.Quantity = CDbl(linea.parent_item.qty_ordered)
                Else
                  oOrder.Lines.Quantity = CDbl(linea.qty_ordered)
                End If
                'JASF (10/08/2022): El descuento a nivel de l�nea ser�: discount_ammount - discount_tax_compensation_amount
                Trace.WriteLineIf(ots.TraceInfo, "discount ammount:" & linea.discount_amount.ToString)
                Trace.WriteLineIf(ots.TraceInfo, "discount_tax_compensation_amount:" & linea.discount_tax_compensation_amount.ToString)
                Trace.WriteLineIf(ots.TraceInfo, "linetotal:" & oOrder.Lines.LineTotal.ToString)
                Trace.WriteLineIf(ots.TraceInfo, "discount %:" & linea.discount_amount * 100 / (oOrder.Lines.LineTotal + linea.discount_amount))
                ''Dim DiscountAmount As Double = linea.discount_amount
                ''Dim DiscountTaxParentAmount As Double = linea.discount_tax_compensation_amount
                ''Dim Descuento As Double = DiscountAmount '- DiscountTaxParentAmount
                'oOrder.Lines.DiscountPercent = linea.discount_amount * 100 / (oOrder.Lines.LineTotal + linea.discount_amount) 'Descuento

                If sStoreName.StartsWith("B2C") Then
                  oOrder.Lines.CostingCode = "E-com002"
                Else
                  oOrder.Lines.CostingCode = "HQ-003"
                End If
                If Not IsNothing(linea.extension_attributes) Then
                  If linea.extension_attributes.delivery_date_item.ToShortDateString <> "01/01/0001" Then
                    oOrder.Lines.UserFields.Fields.Item("U_ART_DeliveryDate").Value = linea.extension_attributes.delivery_date_item
                  End If
                  If Not IsNothing(linea.extension_attributes.pgw) Then
                    If linea.extension_attributes.pgw.requested_delivery_end_date.ToShortDateString <> "01/01/0001" Then oOrder.Lines.ShipDate = linea.extension_attributes.pgw.requested_delivery_end_date
                    If Not IsNothing(linea.extension_attributes.pgw.comment) Then oOrder.Lines.FreeText = linea.extension_attributes.pgw.comment
                  End If
                Else
                  oOrder.Lines.ShipDate = Now.Date
                End If
                'El DocDueDate del documento debe ser el mayor de las fechas de las l�neas
                If oOrder.Lines.ShipDate > DeliveryDate Then
                  DeliveryDate = oOrder.Lines.ShipDate
                End If
              Next

              'Asignamos el ALmac�n a las l�neas del pedido:
              ASSIGN_WAREHOUSE_TO_ORDER(oOrder)

              oOrder.DocDueDate = DeliveryDate
              'Agregamos el pedido
              Trace.WriteLineIf(ots.TraceInfo, "Adding Sales Order...")

              If Not oCompany.InTransaction Then
                oCompany.StartTransaction()
              End If

              If oOrder.Add <> 0 Then Throw New Exception("Error adding Sales Order " & Documento & " in SAP: " & oCompany.GetLastErrorDescription)
              If Not IsNothing(oOrder) Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oOrder)
              If Not IsNothing(oRs) Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
              oOrder = Nothing
              oRs = Nothing

              'Marcar el pedido en Magento como Sincronizado
              Trace.WriteLineIf(ots.TraceInfo, "Updating Sales Order's status in MAGENTO")
              webRequest = HttpWebRequest.Create(URL + "V1/sapSync/" & Documento)
              webRequest.AllowAutoRedirect = True
              webRequest.Timeout = -1
              webRequest.Method = WebRequestMethods.Http.Post
              webRequest.ContentType = "application/json"
              webRequest.Headers.Set("Authorization", "Bearer " & Token)

              webResponse = webRequest.GetResponse()
              If webResponse.StatusCode <> HttpStatusCode.OK Then
                Throw New Exception("Error updating Sales Order's status in MAGENTO")
              End If
              Trace.WriteLineIf(ots.TraceInfo, "Sales Order generated correctly in SAP with DocEntry : " & oCompany.GetNewObjectKey)

            Else
              webResponse.Close()
              Throw New Exception("Error getting Sales Order's info: " & Documento)
            End If
            webResponse.Close()
            If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
          Catch ex As Exception
            If Not IsNothing(oOrder) Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oOrder)
            oOrder = Nothing
            If Not IsNothing(oRs) Then System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
            oRs = Nothing
            Trace.WriteLineIf(ots.TraceError, ex.Message)
            Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
            If oCompany.InTransaction Then
              oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            RECORD_ERRORS(ex.Message, "GENERATE SALES ORDER FROM MAGENTO")
          End Try

        Next
      Else
        Trace.WriteLineIf(ots.TraceInfo, "There are no sales orders pending to sincronize from Magento")
      End If
      Trace.WriteLineIf(ots.TraceInfo, "--------------------------PROCESS END---------------------------------")
      oResponse = New Helpers.DownloadSalesOrdersResponse With {
            .ExecutionSuccess = True,
            .FailureReason = String.Empty
            }
    Catch ex As Exception
      Trace.WriteLineIf(ots.TraceError, ex.Message)
      Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
      RECORD_ERRORS(ex.Message, "GENERATE SALES ORDER FROM MAGENTO")
      oResponse = New Helpers.DownloadSalesOrdersResponse With {
            .ExecutionSuccess = False,
            .FailureReason = ex.Message
            }
    End Try
    Return oResponse

  End Function

#End Region

#Region "Sincronizaci�n de Art�culos (SAP a Magento)"

  Public Function UPDATEItems_NOSYNC() As Helpers.UpdateItemsResponse
    Dim oResponse As Helpers.UpdateItemsResponse = Nothing
    Try
      oResponse = UPDATEItemsEX_NOSYNC()
    Catch ex As Exception
      Trace.WriteLineIf(ots.TraceError, ex.Message)
      Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
      oResponse = New Helpers.UpdateItemsResponse With {
            .ExecutionSuccess = False,
            .FailureReason = ex.Message
            }
    End Try
    Return oResponse
  End Function

  Private Function UPDATEItemsEX_NOSYNC() As Helpers.UpdateItemsResponse
    Dim oResponse As Helpers.UpdateItemsResponse = Nothing
    Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
    Dim oRecordSetUpdate As SAPbobsCOM.Recordset = Nothing
    Try
      Dim URL As String = vbNullString
      Dim Token As String = vbNullString
      oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
      'Buscamos la URL en la tabla de configuraci�n de la integraci�n ART_INTEGRATION
      Dim QueryURL As String = vbNullString
      Select Case oCompany.DbServerType
        Case SAPbobsCOM.BoDataServerTypes.dst_HANADB
          QueryURL = "SELECT T0.""Name"" FROM ""@ART_CONFIGURATIONS""  T0 WHERE T0.""Code"" ='URL'"
        Case Else
          QueryURL = "SELECT T0.Name FROM [@ART_CONFIGURATIONS]  T0 WHERE T0.Code ='URL'"
      End Select
      oRecordSet.DoQuery(QueryURL)
      If Not IsNothing(oRecordSet) AndAlso oRecordSet.RecordCount > 0 Then
        URL = oRecordSet.Fields.Item(0).Value.ToString.Trim
      Else
        Throw New Exception("No est� definida la URL de conexi�n con MAGENTO")
      End If
      If Not URL.EndsWith("/") Then URL = URL & "/"

      'Buscamos el BearerToken para conexi�n
      Dim QueryTOKEN As String = vbNullString
      Select Case oCompany.DbServerType
        Case SAPbobsCOM.BoDataServerTypes.dst_HANADB
          QueryTOKEN = "SELECT T0.""Name"" FROM ""@ART_CONFIGURATIONS""  T0 WHERE T0.""Code"" ='TOKEN'"
        Case Else
          QueryTOKEN = "SELECT T0.Name FROM [@ART_CONFIGURATIONS] T0 WHERE T0.Code = 'TOKEN'"
      End Select
      oRecordSet.DoQuery(QueryTOKEN)
      If Not IsNothing(oRecordSet) AndAlso oRecordSet.RecordCount > 0 Then
        Token = oRecordSet.Fields.Item(0).Value.ToString.Trim
      Else
        Throw New Exception("No est� definido el TOKEN de conexi�n con MAGENTO")
      End If
      'rpaya(20220330)
      'Dim Query As String = "SELECT TOP 1 T1.""ItemCode"", T0.""ItemName"", T1.""Price"", CASE WHEN T0.""QryGroup50"" = 'Y' THEN 'cern' 
      '                        WHEN T0.""QryGroup51"" = 'Y' THEN 'b2c' WHEN ""QryGroup52"" = 'Y' THEN 'b2b' ELSE '' END 
      '                        FROM OITM T0  INNER JOIN ITM1 T1 ON T0.""ItemCode"" = T1.""ItemCode"" WHERE T1.""PriceList""  = 19 AND  T0.""U_ART_MagentoSync"" ='N' 
      '                        AND  T0.""validFor"" ='Y' ORDER BY T1.""Price"" DESC"
      'Dim Query As String = "SELECT TOP 10 T1.""ItemCode"", T0.""ItemName"", T1.""Price"",T0.""QryGroup50"",T0.""QryGroup51"",""T0"".""QryGroup52"" 
      '                        FROM OITM T0  INNER JOIN ITM1 T1 ON T0.""ItemCode"" = T1.""ItemCode"" WHERE T1.""PriceList""  = 19 AND  T0.""U_ART_MagentoSync"" ='N' 
      '                        AND  T0.""validFor"" ='Y' ORDER BY T1.""Price"" DESC"
      Dim Query As String = "Select * From ART_MAGENTO_ITEM_AU_VIEW"
      oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
      oRecordSet.DoQuery(Query)
      If oRecordSet.RecordCount > 0 Then
        oRecordSet.MoveFirst()
        Dim webRequest As HttpWebRequest
        'rpaya(20220329)
        'Dim strURL As String = URL & "default/async/bulk/V1/products"
        Dim strURL As String = URL & "V1/products"
        Dim Items As New List(Of Wrapper_ItemProduct)
        While Not oRecordSet.EoF
          Dim ItemDetail As New Wrapper_ItemProduct
          ItemDetail.sku = oRecordSet.Fields.Item("ItemCode").Value.ToString
          ItemDetail.name = oRecordSet.Fields.Item("ItemName").Value.ToString
          ItemDetail.price = CDbl(oRecordSet.Fields.Item("Price").Value)
          ItemDetail.status = 1
          '***********************************************************
          'Asignaci�n de WebSites
          '************************************************************
          '1) Si la propiedad 50 est� marcada se enviar� el valor 1
          '2) Si la propiedad 51 est� marcada se enviar� el valor 19
          '3) Si la propiedad 51 est� marcada se enviar� el valor 23
          '***********************************************************
          Dim ListWebSiteIds As New List(Of Integer)

          If oRecordSet.Fields.Item("QryGroup50").Value.ToString.Equals("Y") Or oRecordSet.Fields.Item("QryGroup51").Value.ToString.Equals("Y") Or oRecordSet.Fields.Item("QryGroup52").Value.ToString.Equals("Y") Then
            If ListWebSiteIds.Count > 0 Then ListWebSiteIds.Clear()
            If oRecordSet.Fields.Item("QryGroup50").Value.ToString.Equals("Y") Then ListWebSiteIds.Add(1)
            If oRecordSet.Fields.Item("QryGroup51").Value.ToString.Equals("Y") Then ListWebSiteIds.Add(19)
            If oRecordSet.Fields.Item("QryGroup52").Value.ToString.Equals("Y") Then ListWebSiteIds.Add(23)
            ItemDetail.extension_attributes.website_ids = ListWebSiteIds

          End If

          'rpaya(20220329)
          'Items.Add(ItemDetail)
          '**************************************************************************************************************************************
          'Actualizar el art�culo
          '*************************************************************************************************************************************
          ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf AcceptAllCertifications)
          ServicePointManager.Expect100Continue = True
          ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
          Dim ServiceCall As String = JsonConvert.SerializeObject(ItemDetail)
          ServiceCall = "{ ""product"": " & ServiceCall & "}"
          Trace.WriteLineIf(ots.TraceInfo, "Call to Magento: " & ServiceCall)
          Dim data = Encoding.UTF8.GetBytes(ServiceCall)
          webRequest = HttpWebRequest.Create(strURL)
          webRequest.AllowAutoRedirect = True
          webRequest.Timeout = -1
          webRequest.Method = WebRequestMethods.Http.Post
          webRequest.ContentLength = data.Length
          webRequest.ContentType = "application/json"
          webRequest.Headers.Set("Authorization", "Bearer " & Token)
          Dim stream = webRequest.GetRequestStream()
          stream.Write(data, 0, data.Length)
          stream.Close()
          stream.Dispose()
          Try
            Dim webResponse As HttpWebResponse = webRequest.GetResponse()

            If webResponse.StatusCode = HttpStatusCode.OK Then

              Trace.WriteLineIf(ots.TraceInfo, "Item with SAP Code: " & oRecordSet.Fields.Item("ItemCode").Value.ToString.Trim & " updated to MAGENTO")
              Dim receiveStream As Stream = webResponse.GetResponseStream
              Dim readStream As StreamReader = New StreamReader(receiveStream, Encoding.UTF8)
              Dim respuesta As String = readStream.ReadToEnd
              Dim sCorreo As String = String.Empty
              'Dim Resp = JsonConvert.DeserializeObject(respuesta)
              'Dim Billing_street As String = Resp("billing_address")("street").ToString.Replace("[", "").Replace("]", "")
              'Dim Shipping_street As String = Resp("extension_attributes")("shipping_assignments")(0)("shipping")("address")("street").ToString.Replace("[", "").Replace("]", "")
              Dim oitem As New Wrapper_ItemRes
              oitem = JsonConvert.DeserializeObject(Of Wrapper_ItemRes)(respuesta)
              Trace.WriteLineIf(ots.TraceInfo, "MAGENTO Answer: " & respuesta)

              Dim QueryUpdate As String = "UPDATE OITM SET ""U_ART_MagentoSync"" ='Y' WHERE ""ItemCode"" = '" & oRecordSet.Fields.Item("ItemCode").Value.ToString.Trim & "'"
              Dim oRs As SAPbobsCOM.Recordset
              oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
              'webResponse.Heade
              oRs.DoQuery(QueryUpdate)
              '220715
              QueryUpdate = "UPDATE OITM SET ""U_WebID"" = '" & oitem.id & "' WHERE ""ItemCode"" = '" & oRecordSet.Fields.Item("ItemCode").Value.ToString.Trim & "'"
              oRs.DoQuery(QueryUpdate)
              System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
            End If


            webResponse.Close()
          Catch ex As Exception
            RECORD_ERRORS(ex.Message, "UPDATE ITEMS")
            Trace.WriteLineIf(ots.TraceError, ex.Message)
            Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
          Finally
          End Try
          stream = Nothing

          '*********************************************************************************************************************************************************************************************
          'Si un art�culo no tiene indicada ninguna propiedad, por defecto, se crea con la propiedad 50 ("cern") marcada siendo este comportamiento incorrecto, lo que vamos a hacer es lo siguiente:
          ' Una vez se cree el art�culo debemos desmarcar la propiedad.
          '*********************************************************************************************************************************************************************************************
          If ListWebSiteIds.Count = 0 Then

            Dim webRequest2 As HttpWebRequest
            Dim strURL2 As String = URL & "V1/products/" & oRecordSet.Fields.Item("ItemCode").Value.ToString.Trim & "/websites/1"
            ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf AcceptAllCertifications)
            ServicePointManager.Expect100Continue = True
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            webRequest2 = HttpWebRequest.Create(strURL2)
            webRequest2.AllowAutoRedirect = True
            webRequest2.Timeout = -1
            webRequest2.Method = "DELETE"
            webRequest2.ContentType = "application/json"
            webRequest2.Headers.Set("Authorization", "Bearer " & Token)
            Dim stream2 = webRequest2.GetRequestStream()
            stream2.Close()
            stream2.Dispose()
            Try
              Dim webResponse2 As HttpWebResponse = webRequest2.GetResponse()
              If webResponse2.StatusCode = HttpStatusCode.OK Then

                Trace.WriteLineIf(ots.TraceInfo, "Item with SAP Code: " & oRecordSet.Fields.Item("ItemCode").Value.ToString.Trim & ". Property 50(cern) unmarked with success in MAGENTO")
              End If
              webResponse2.Close()
            Catch ex As Exception
              RECORD_ERRORS(ex.Message, "UPDATE ITEMS")
              Trace.WriteLineIf(ots.TraceError, ex.Message)
              Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
            Finally
            End Try
            stream2 = Nothing

          End If


          oRecordSet.MoveNext()
        End While
        'rpaya(20220329)
        'Dim Item As New Wrapper_Items
        'Item.product = Items
        '    'Llamamos al SW
        'rpaya(20220329)
        '***********************************************************************************************************************************
        'Muevo este trozo de c�digo a otra parte porque se van a actualizar los art�culos uno a uno.
        '************************************************************************************************************************************
        'ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf AcceptAllCertifications)
        'ServicePointManager.Expect100Continue = True
        'ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        'Dim ServiceCall As String = JsonConvert.SerializeObject(Item)
        'ServiceCall = "[" & ServiceCall & "]"
        'Trace.WriteLineIf(ots.TraceInfo, "Call to Magento: " & ServiceCall)
        'Dim data = Encoding.UTF8.GetBytes(ServiceCall)
        'webRequest = HttpWebRequest.Create(strURL)
        'webRequest.AllowAutoRedirect = True
        'webRequest.Timeout = -1
        'webRequest.Method = WebRequestMethods.Http.Post
        'webRequest.ContentLength = data.Length
        'webRequest.ContentType = "application/json"
        'webRequest.Headers.Set("Authorization", "Bearer " & Token)
        'Dim stream = webRequest.GetRequestStream()
        'stream.Write(data, 0, data.Length)
        'stream.Close()
        'stream.Dispose()
        'Try
        '    Dim webResponse As HttpWebResponse = webRequest.GetResponse()
        '    If webResponse.StatusCode = HttpStatusCode.OK Then

        '        Trace.WriteLineIf(ots.TraceInfo, "Item with SAP Code: " & oRecordSet.Fields.Item("ItemCode").Value.ToString.Trim & " updated to MAGENTO")
        '    End If
        '    Dim oItem As SAPbobsCOM.Items
        '    oItem = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        '    oItem.GetByKey(oRecordSet.Fields.Item("ItemCode").Value.ToString.Trim)
        '    oItem.UserFields.Fields.Item("U_ART_MagentoSync").Value = "Y"
        '    oItem.Update()
        '    webResponse.Close()
        'Catch ex As Exception
        '    RECORD_ERRORS(ex.Message, "UPDATE ITEMS")
        '    Trace.WriteLineIf(ots.TraceError, ex.Message)
        '    Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
        'Finally
        'End Try
        'stream = Nothing
      Else
        Trace.WriteLineIf(ots.TraceInfo, "There are no ITEMS needed to sync with Magento")
      End If
      oResponse = New Helpers.UpdateItemsResponse With {
                        .ExecutionSuccess = True,
                        .FailureReason = String.Empty
                        }
    Catch ex As Exception
      Trace.WriteLineIf(ots.TraceError, ex.Message)
      Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
      oResponse = New Helpers.UpdateItemsResponse With {
            .ExecutionSuccess = False,
            .FailureReason = ex.Message
            }

    End Try
    Return oResponse

  End Function

#End Region

#Region "Sincronizaci�n de Clientes (SAP a Magento)"

  Public Function UPDATEclients_NOSYNC() As Helpers.UpdateClientsResponse
    Dim oResponse As Helpers.UpdateClientsResponse = Nothing
    Try
      oResponse = UPDATEClientsEX_NOSYNC()
    Catch ex As Exception
      Trace.WriteLineIf(ots.TraceError, ex.Message)
      Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
      oResponse = New Helpers.UpdateClientsResponse With {
                  .ExecutionSuccess = False,
                  .FailureReason = ex.Message
                  }
    End Try
    Return oResponse
  End Function

  Private Function UPDATEClientsEX_NOSYNC() As Helpers.UpdateClientsResponse
    Dim oResponse As Helpers.UpdateClientsResponse = Nothing
    Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
    Dim oRecordSetUpdate As SAPbobsCOM.Recordset = Nothing
    Try
      Dim URL As String = vbNullString
      Dim Token As String = vbNullString
      oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
      'Buscamos la URL en la tabla de configuraci�n de la integraci�n ART_INTEGRATION
      Dim QueryURL As String = vbNullString
      Select Case oCompany.DbServerType
        Case SAPbobsCOM.BoDataServerTypes.dst_HANADB
          QueryURL = "SELECT T0.""Name"" FROM ""@ART_CONFIGURATIONS""  T0 WHERE T0.""Code"" ='URL2'"
        Case Else
          QueryURL = "SELECT T0.Name FROM [@ART_CONFIGURATIONS]  T0 WHERE T0.Code ='URL2'"
      End Select
      oRecordSet.DoQuery(QueryURL)
      If Not IsNothing(oRecordSet) AndAlso oRecordSet.RecordCount > 0 Then
        URL = oRecordSet.Fields.Item(0).Value.ToString.Trim
      Else
        Throw New Exception("No est� definida la URL de conexi�n con MAGENTO")
      End If
      If Not URL.EndsWith("/") Then URL = URL & "/"

      'Buscamos el BearerToken para conexi�n
      Dim QueryTOKEN As String = vbNullString
      Select Case oCompany.DbServerType
        Case SAPbobsCOM.BoDataServerTypes.dst_HANADB
          QueryTOKEN = "SELECT T0.""Name"" FROM ""@ART_CONFIGURATIONS""  T0 WHERE T0.""Code"" ='TOKEN'"
        Case Else
          QueryTOKEN = "SELECT T0.Name FROM [@ART_CONFIGURATIONS] T0 WHERE T0.Code = 'TOKEN'"
      End Select
      oRecordSet.DoQuery(QueryTOKEN)
      If Not IsNothing(oRecordSet) AndAlso oRecordSet.RecordCount > 0 Then
        Token = oRecordSet.Fields.Item(0).Value.ToString.Trim
      Else
        Throw New Exception("No est� definido el TOKEN de conexi�n con MAGENTO")
      End If

      Dim correoant As String = ""

      Dim Query As String = "select * from ART_MAGENTO_CUSTOMER_AU_VIEW"
      oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
      oRecordSet.DoQuery(Query)
      If oRecordSet.RecordCount > 0 Then
        oRecordSet.MoveFirst()
        Dim webRequest As HttpWebRequest
        'rpaya(20220329)
        'Dim strURL As String = URL & "default/async/bulk/V1/products"
        Dim strURL As String = URL & "V1/customers"
        'Dim customers As New List(Of Wrapper_Client)
        Dim customer As New Wrapper_Client
        Dim dir As New Wrapper_ClientAddress
        Dim reg As New Wrapper_ClientRegion
        Dim ServiceCall As String
        Dim data As Byte()
        Dim stream As Stream
        Dim pass As String = String.Empty
        Dim card As String = String.Empty
        Dim email As String = String.Empty
        While Not oRecordSet.EoF

          If correoant <> oRecordSet.Fields.Item("email").Value.ToString Then
            If correoant <> "" Then
              '**************************************************************************************************************************************
              'Actualizar el cliente
              '*************************************************************************************************************************************
              ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf AcceptAllCertifications)
              ServicePointManager.Expect100Continue = True
              ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
              ServiceCall = JsonConvert.SerializeObject(customer)
              'ServiceCall = "{ ""customer"": " & ServiceCall.Replace("email"":", "email"":""").Replace(",""first", """,""first") & ",""password"":""" & oRecordSet.Fields.Item("password").Value.ToString & """}"
              ServiceCall = "{ ""customer"": " & ServiceCall & " ,""password"":""" & pass & """ }"
              Trace.WriteLineIf(ots.TraceInfo, "URL: " & strURL)
              Trace.WriteLineIf(ots.TraceInfo, "Call to Magento: " & ServiceCall)
              data = Encoding.UTF8.GetBytes(ServiceCall)
              webRequest = HttpWebRequest.Create(strURL)
              webRequest.AllowAutoRedirect = True
              webRequest.Timeout = -1
              webRequest.Method = WebRequestMethods.Http.Post
              webRequest.ContentLength = data.Length
              webRequest.ContentType = "application/json"
              webRequest.Headers.Set("Authorization", "Bearer " & Token)
              stream = webRequest.GetRequestStream()
              stream.Write(data, 0, data.Length)
              stream.Close()
              stream.Dispose()
              Try
                Dim webResponse As HttpWebResponse = webRequest.GetResponse()
                If webResponse.StatusCode = HttpStatusCode.OK Then
                  Dim receiveStream As Stream = webResponse.GetResponseStream
                  Dim readStream As StreamReader = New StreamReader(receiveStream, Encoding.UTF8)
                  Dim respuesta As String = readStream.ReadToEnd
                  Dim sCorreo As String = String.Empty

                  'Dim Resp = JsonConvert.DeserializeObject(respuesta)
                  'Dim Billing_street As String = Resp("billing_address")("street").ToString.Replace("[", "").Replace("]", "")
                  'Dim Shipping_street As String = Resp("extension_attributes")("shipping_assignments")(0)("shipping")("address")("street").ToString.Replace("[", "").Replace("]", "")
                  Dim ouser As New Wrapper_ClientRes
                  ouser = JsonConvert.DeserializeObject(Of Wrapper_ClientRes)(respuesta)
                  Trace.WriteLineIf(ots.TraceInfo, "MAGENTO Answer: " & respuesta)

                  Trace.WriteLineIf(ots.TraceInfo, "Customer with SAP Code: " & card & " updated to MAGENTO")
                  Dim QueryUpdate As String = "UPDATE OCPR SET ""U_magentoLoginId"" = '" & ouser.id & "' WHERE ""CardCode"" = '" & card & "' and ""E_MailL"" = '" & email & "'"
                  Dim oRs As SAPbobsCOM.Recordset
                  oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                  oRs.DoQuery(QueryUpdate)

                  QueryUpdate = "UPDATE OCRD SET ""U_ART_MagentoSync"" = 'Y' WHERE ""CardCode"" = '" & card & "'"
                  oRs.DoQuery(QueryUpdate)

                  System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
                  'Else
                  '  Throw New Exception(webResponse.StatusCode & "-" & webResponse.StatusDescription)
                End If
                'Dim QueryUpdate As String = "UPDATE OCRD SET ""U_ART_MagentoSync"" = 'Y' WHERE ""CardCode"" = '" & oRecordSet.Fields.Item("CardCode").Value.ToString.Trim & "'"



                webResponse.Close()
              Catch ex As Exception
                RECORD_ERRORS(ex.Message, "UPDATE Customers")
                Trace.WriteLineIf(ots.TraceError, ex.Message)
                Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
              Finally
              End Try
              stream = Nothing

            End If

            customer = New Wrapper_Client
            Trace.WriteLineIf(ots.TraceInfo, "customer.id: " & oRecordSet.Fields.Item("id").Value.ToString)
            'If oRecordSet.Fields.Item("id").Value.ToString <> "0" Then
            customer.id = oRecordSet.Fields.Item("id").Value.ToString
            Trace.WriteLineIf(ots.TraceInfo, "customer.group_id: " & oRecordSet.Fields.Item("group_id").Value.ToString)
            customer.group_id = oRecordSet.Fields.Item("group_id").Value.ToString
            Trace.WriteLineIf(ots.TraceInfo, "customer.email: " & oRecordSet.Fields.Item("email").Value.ToString)
            customer.email = oRecordSet.Fields.Item("email").Value.ToString
            Trace.WriteLineIf(ots.TraceInfo, "customer.firstname: " & oRecordSet.Fields.Item("firstname").Value.ToString)
            customer.firstname = oRecordSet.Fields.Item("firstname").Value.ToString
            Trace.WriteLineIf(ots.TraceInfo, "customer.lastname: " & oRecordSet.Fields.Item("lastname").Value.ToString)
            customer.lastname = oRecordSet.Fields.Item("lastname").Value.ToString
            Trace.WriteLineIf(ots.TraceInfo, "customer.website_id: " & oRecordSet.Fields.Item("website_id").Value.ToString)
            customer.website_id = oRecordSet.Fields.Item("website_id").Value.ToString

            'Trace.WriteLineIf(ots.TraceInfo, "customer.sendemail_store_id: " & oRecordSet.Fields.Item("sendemail_store_id").Value.ToString)

            pass = oRecordSet.Fields.Item("password").Value.ToString
            card = oRecordSet.Fields.Item("CardCode").Value.ToString
            email = oRecordSet.Fields.Item("email").Value.ToString

            dir.id = oRecordSet.Fields.Item(8).Value.ToString
            reg.region_id = oRecordSet.Fields.Item(11).Value.ToString
            reg.region = oRecordSet.Fields.Item(10).Value.ToString
            reg.region_code = oRecordSet.Fields.Item(9).Value.ToString
            dir.region = reg
            dir.region_id = oRecordSet.Fields.Item(12).Value.ToString
            dir.street.Add(oRecordSet.Fields.Item(13).Value.ToString)
            dir.company = oRecordSet.Fields.Item(14).Value.ToString
            dir.telephone = oRecordSet.Fields.Item(15).Value.ToString
            dir.fax = oRecordSet.Fields.Item(16).Value.ToString
            dir.postcode = oRecordSet.Fields.Item(17).Value.ToString
            dir.city = oRecordSet.Fields.Item(18).Value.ToString
            dir.firstname = oRecordSet.Fields.Item(19).Value.ToString
            dir.default_shipping = oRecordSet.Fields.Item(20).Value.ToString
            dir.default_billing = oRecordSet.Fields.Item(21).Value.ToString

            customer.addresses.Add(dir)

          Else
            dir.id = oRecordSet.Fields.Item(8).Value.ToString
            reg.region_id = oRecordSet.Fields.Item(11).Value.ToString
            reg.region = oRecordSet.Fields.Item(10).Value.ToString
            reg.region_code = oRecordSet.Fields.Item(9).Value.ToString
            dir.region = reg
            dir.region_id = oRecordSet.Fields.Item(12).Value.ToString
            dir.street.Add(oRecordSet.Fields.Item(13).Value.ToString)
            dir.company = oRecordSet.Fields.Item(14).Value.ToString
            dir.telephone = oRecordSet.Fields.Item(15).Value.ToString
            dir.fax = oRecordSet.Fields.Item(16).Value.ToString
            dir.postcode = oRecordSet.Fields.Item(17).Value.ToString
            dir.city = oRecordSet.Fields.Item(18).Value.ToString
            dir.firstname = oRecordSet.Fields.Item(19).Value.ToString
            dir.default_shipping = oRecordSet.Fields.Item(20).Value.ToString
            dir.default_billing = oRecordSet.Fields.Item(21).Value.ToString

            customer.addresses.Add(dir)

          End If



          correoant = oRecordSet.Fields.Item("email").Value.ToString
          oRecordSet.MoveNext()
        End While
        '**************************************************************************************************************************************
        'Actualizar el cliente
        '*************************************************************************************************************************************
        ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf AcceptAllCertifications)
        ServicePointManager.Expect100Continue = True
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        ServiceCall = JsonConvert.SerializeObject(customer)
        'ServiceCall = "{ ""customer"": " & ServiceCall.Replace("email"":", "email"":""").Replace(",""first", """,""first") & ",""password"":""" & oRecordSet.Fields.Item("password").Value.ToString & """}"
        ServiceCall = "{ ""customer"": " & ServiceCall & " ,""password"":""" & pass & """ }"
        Trace.WriteLineIf(ots.TraceInfo, "Call to Magento: " & ServiceCall)
        data = Encoding.UTF8.GetBytes(ServiceCall)
        webRequest = HttpWebRequest.Create(strURL)
        webRequest.AllowAutoRedirect = True
        webRequest.Timeout = -1
        webRequest.Method = WebRequestMethods.Http.Post
        webRequest.ContentLength = data.Length
        webRequest.ContentType = "application/json"
        webRequest.Headers.Set("Authorization", "Bearer " & Token)
        'webRequest.Headers.Set("Cookie", "PHPSESSID=be901b47955ad309a97f392383963c65; mage-messages=%5B%7B%22type%22%3A%22error%22%2C%22text%22%3A%22Invalid%20Form%20Key.%20Please%20refresh%20the%20page.%22%7D%2C%7B%22type%22%3A%22error%22%2C%22text%22%3A%22Invalid%20Form%20Key.%20Please%20refresh%20the%20page.%22%7D%2C%7B%22type%22%3A%22error%22%2C%22text%22%3A%22Invalid%20Form%20Key.%20Please%20refresh%20the%20page.%22%7D%2C%7B%22type%22%3A%22error%22%2C%22text%22%3A%22Invalid%20Form%20Key.%20Please%20refresh%20the%20page.%22%7D%2C%7B%22type%22%3A%22error%22%2C%22text%22%3A%22Invalid%20Form%20Key.%20Please%20refresh%20the%20page.%22%7D%2C%7B%22type%22%3A%22error%22%2C%22text%22%3A%22Invalid%20Form%20Key.%20Please%20refresh%20the%20page.%22%7D%2C%7B%22type%22%3A%22error%22%2C%22text%22%3A%22Invalid%20Form%20Key.%20Please%20refresh%20the%20page.%22%7D%2C%7B%22type%22%3A%22error%22%2C%22text%22%3A%22Invalid%20Form%20Key.%20Please%20refresh%20the%20page.%22%7D%2C%7B%22type%22%3A%22error%22%2C%22text%22%3A%22Invalid%20Form%20Key.%20Please%20refresh%20the%20page.%22%7D%2C%7B%22type%22%3A%22error%22%2C%22text%22%3A%22Invalid%20Form%20Key.%20Please%20refresh%20the%20page.%22%7D%2C%7B%22type%22%3A%22error%22%2C%22text%22%3A%22Invalid%20Form%20Key.%20Please%20refresh%20the%20page.%22%7D%2C%7B%22type%22%3A%22error%22%2C%22text%22%3A%22Invalid%20Form%20Key.%20Please%20refresh%20the%20page.%22%7D%2C%7B%22type%22%3A%22error%22%2C%22text%22%3A%22Invalid%20Form%20Key.%20Please%20refresh%20the%20page.%22%7D%2C%7B%22type%22%3A%22error%22%2C%22text%22%3A%22Invalid%20Form%20Key.%20Please%20refresh%20the%20page.%22%7D%2C%7B%22type%22%3A%22error%22%2C%22text%22%3A%22Invalid%20Form%20Key.%20Please%20refresh%20the%20page.%22%7D%2C%7B%22type%22%3A%22error%22%2C%22text%22%3A%22Invalid%20Form%20Key.%20Please%20refresh%20the%20page.%22%7D%2C%7B%22type%22%3A%22error%22%2C%22text%22%3A%22Invalid%20Form%20Key.%20Please%20refresh%20the%20page.%22%7D%5D; private_content_version=c1845b8807f05d6c96d6476f59c732c9")
        stream = webRequest.GetRequestStream()
        stream.Write(data, 0, data.Length)
        stream.Close()
        stream.Dispose()
        Try
          Dim webResponse As HttpWebResponse = webRequest.GetResponse()
          If webResponse.StatusCode = HttpStatusCode.OK Then
            Dim receiveStream As Stream = webResponse.GetResponseStream
            Dim readStream As StreamReader = New StreamReader(receiveStream, Encoding.UTF8)
            Dim respuesta As String = readStream.ReadToEnd
            Dim sCorreo As String = String.Empty
            'Dim Resp = JsonConvert.DeserializeObject(respuesta)
            'Dim Billing_street As String = Resp("billing_address")("street").ToString.Replace("[", "").Replace("]", "")
            'Dim Shipping_street As String = Resp("extension_attributes")("shipping_assignments")(0)("shipping")("address")("street").ToString.Replace("[", "").Replace("]", "")
            Dim ouser As New Wrapper_ClientRes
            ouser = JsonConvert.DeserializeObject(Of Wrapper_ClientRes)(respuesta)
            Trace.WriteLineIf(ots.TraceInfo, "MAGENTO Answer: " & respuesta)

            Trace.WriteLineIf(ots.TraceInfo, "Customer with SAP Code: " & card & " updated to MAGENTO")
            Dim QueryUpdate As String = "UPDATE OCPR SET ""U_magentoLoginId"" = '" & ouser.id & "' WHERE ""CardCode"" = '" & card & "' and ""E_MailL"" = '" & email & "'"
            Dim oRs As SAPbobsCOM.Recordset
            oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oRs.DoQuery(QueryUpdate)
            QueryUpdate = "UPDATE OCRD SET ""U_ART_MagentoSync"" = 'Y' WHERE ""CardCode"" = '" & card & "'"
            oRs.DoQuery(QueryUpdate)

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
            'Else
            '  Throw New Exception(webResponse.StatusCode & "-" & webResponse.StatusDescription)
          End If
          'Dim QueryUpdate As String = "UPDATE OCRD SET ""U_ART_MagentoSync"" = 'Y' WHERE ""CardCode"" = '" & oRecordSet.Fields.Item("CardCode").Value.ToString.Trim & "'"



          webResponse.Close()
        Catch ex As Exception
          RECORD_ERRORS(ex.Message, "UPDATE Customers")
          Trace.WriteLineIf(ots.TraceError, ex.Message)
          Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
        Finally
        End Try
        stream = Nothing



      Else
        Trace.WriteLineIf(ots.TraceInfo, "There are no Clients needed to sync with Magento")
      End If
      oResponse = New Helpers.UpdateClientsResponse With {
                              .ExecutionSuccess = True,
                              .FailureReason = String.Empty
                              }
    Catch ex As Exception
      Trace.WriteLineIf(ots.TraceError, ex.Message)
      Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
      oResponse = New Helpers.UpdateClientsResponse With {
                  .ExecutionSuccess = False,
                  .FailureReason = ex.Message
                  }

    End Try
    Return oResponse

  End Function

#End Region

#Region "Sincronizaci�n de Precios (SAP a Magento)"

  Public Function UPDATEprices_NOSYNC() As Helpers.UpdatePricesResponse
    Dim oResponse As Helpers.UpdatePricesResponse = Nothing
    Try
      oResponse = UPDATEPricesEX_NOSYNC()
    Catch ex As Exception
      Trace.WriteLineIf(ots.TraceError, ex.Message)
      Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
      oResponse = New Helpers.UpdatePricesResponse With {
                  .ExecutionSuccess = False,
                  .FailureReason = ex.Message
                  }
    End Try
    Return oResponse
  End Function

  Private Function UPDATEPricesEX_NOSYNC() As Helpers.UpdatePricesResponse
    Dim oResponse As Helpers.UpdatePricesResponse = Nothing
    Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
    Dim oRecordSetUpdate As SAPbobsCOM.Recordset = Nothing
    Try
      Dim URL As String = vbNullString
      Dim Token As String = vbNullString
      oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
      'Buscamos la URL en la tabla de configuraci�n de la integraci�n ART_INTEGRATION
      Dim QueryURL As String = vbNullString
      Select Case oCompany.DbServerType
        Case SAPbobsCOM.BoDataServerTypes.dst_HANADB
          QueryURL = "SELECT T0.""Name"" FROM ""@ART_CONFIGURATIONS""  T0 WHERE T0.""Code"" ='URL2'"
        Case Else
          QueryURL = "SELECT T0.Name FROM [@ART_CONFIGURATIONS]  T0 WHERE T0.Code ='URL2'"
      End Select
      oRecordSet.DoQuery(QueryURL)
      If Not IsNothing(oRecordSet) AndAlso oRecordSet.RecordCount > 0 Then
        URL = oRecordSet.Fields.Item(0).Value.ToString.Trim
      Else
        Throw New Exception("No est� definida la URL de conexi�n con MAGENTO")
      End If
      If Not URL.EndsWith("/") Then URL = URL & "/"

      'Buscamos el BearerToken para conexi�n
      Dim QueryTOKEN As String = vbNullString
      Select Case oCompany.DbServerType
        Case SAPbobsCOM.BoDataServerTypes.dst_HANADB
          QueryTOKEN = "SELECT T0.""Name"" FROM ""@ART_CONFIGURATIONS""  T0 WHERE T0.""Code"" ='TOKEN'"
        Case Else
          QueryTOKEN = "SELECT T0.Name FROM [@ART_CONFIGURATIONS] T0 WHERE T0.Code = 'TOKEN'"
      End Select
      oRecordSet.DoQuery(QueryTOKEN)
      If Not IsNothing(oRecordSet) AndAlso oRecordSet.RecordCount > 0 Then
        Token = oRecordSet.Fields.Item(0).Value.ToString.Trim
      Else
        Throw New Exception("No est� definido el TOKEN de conexi�n con MAGENTO")
      End If

      Dim correoant As String = ""

      'BASE prices
      Dim Query As String = "select * from ART_MAGENTO_BASEPRICES "
      oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
      oRecordSet.DoQuery(Query)
      If oRecordSet.RecordCount > 0 Then
        oRecordSet.MoveFirst()
        Dim webRequest As HttpWebRequest
        'rpaya(20220329)
        'Dim strURL As String = URL & "default/async/bulk/V1/products"
        Dim strURL As String = URL & "V1/products/base-prices"
        'Dim customers As New List(Of Wrapper_Client)
        Dim basepri As New Wrapper_BasePrice


        Dim ServiceCall As String
        Dim data As Byte()
        Dim stream As Stream

        While Not oRecordSet.EoF

          basepri = New Wrapper_BasePrice
          Trace.WriteLineIf(ots.TraceInfo, "basepri.price: " & oRecordSet.Fields.Item(0).Value.ToString.ToString.Replace(",", "."))
          basepri.price = oRecordSet.Fields.Item(0).Value.ToString.Replace(",", ".")
          Trace.WriteLineIf(ots.TraceInfo, "basepri.sku: " & oRecordSet.Fields.Item(2).Value.ToString)
          basepri.sku = oRecordSet.Fields.Item(2).Value.ToString
          Trace.WriteLineIf(ots.TraceInfo, "basepri.store_id: " & oRecordSet.Fields.Item(1).Value.ToString)
          basepri.store_id = oRecordSet.Fields.Item(1).Value.ToString


          '**************************************************************************************************************************************
          'Actualizar el precio base
          '*************************************************************************************************************************************
          ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf AcceptAllCertifications)
          ServicePointManager.Expect100Continue = True
          ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
          ServiceCall = JsonConvert.SerializeObject(basepri)
          'ServiceCall = "{ ""customer"": " & ServiceCall.Replace("email"":", "email"":""").Replace(",""first", """,""first") & ",""password"":""" & oRecordSet.Fields.Item("password").Value.ToString & """}"
          ServiceCall = "{ ""prices"": [" & ServiceCall & "] }"

          Trace.WriteLineIf(ots.TraceInfo, "Call to Magento: " & ServiceCall)
          data = Encoding.UTF8.GetBytes(ServiceCall)
          webRequest = HttpWebRequest.Create(strURL)
          webRequest.AllowAutoRedirect = True
          webRequest.Timeout = -1
          webRequest.Method = WebRequestMethods.Http.Post
          webRequest.ContentLength = data.Length
          webRequest.ContentType = "application/json"
          webRequest.Headers.Set("Authorization", "Bearer " & Token)
          stream = webRequest.GetRequestStream()
          stream.Write(data, 0, data.Length)
          stream.Close()
          stream.Dispose()
          Try
            Dim webResponse As HttpWebResponse = webRequest.GetResponse()
            If webResponse.StatusCode = HttpStatusCode.OK Then
              Dim receiveStream As Stream = webResponse.GetResponseStream
              Dim readStream As StreamReader = New StreamReader(receiveStream, Encoding.UTF8)
              Dim respuesta As String = readStream.ReadToEnd
              Dim sCorreo As String = String.Empty

              'Dim Resp = JsonConvert.DeserializeObject(respuesta)
              'Dim Billing_street As String = Resp("billing_address")("street").ToString.Replace("[", "").Replace("]", "")
              'Dim Shipping_street As String = Resp("extension_attributes")("shipping_assignments")(0)("shipping")("address")("street").ToString.Replace("[", "").Replace("]", "")
              Dim ouser As New Wrapper_ClientRes
              ouser = JsonConvert.DeserializeObject(Of Wrapper_ClientRes)(respuesta)
              Trace.WriteLineIf(ots.TraceInfo, "MAGENTO Answer: " & respuesta)

              Trace.WriteLineIf(ots.TraceInfo, "Price with Item Code: " & oRecordSet.Fields.Item(1).Value.ToString & " updated to MAGENTO")
              Dim QueryUpdate As String = "UPDATE OITM SET ""U_ART_mgtBasePrice"" = '" & basepri.price.ToString & "' WHERE ""ItemCode"" = '" & oRecordSet.Fields.Item(1).Value.ToString & "'"
              Dim oRs As SAPbobsCOM.Recordset
              oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

              oRs.DoQuery(QueryUpdate)



              System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
              'Else
              '  Throw New Exception(webResponse.StatusCode & "-" & webResponse.StatusDescription)
            End If
            'Dim QueryUpdate As String = "UPDATE OCRD SET ""U_ART_MagentoSync"" = 'Y' WHERE ""CardCode"" = '" & oRecordSet.Fields.Item("CardCode").Value.ToString.Trim & "'"



            webResponse.Close()
          Catch ex As Exception
            RECORD_ERRORS(ex.Message, "UPDATE Base prices")
            Trace.WriteLineIf(ots.TraceError, ex.Message)
            Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
          Finally
          End Try
          stream = Nothing

          oRecordSet.MoveNext()
        End While


      Else
        Trace.WriteLineIf(ots.TraceInfo, "There are no Base prices needed to sync with Magento")
      End If

      'TIER prices
      Query = "select * from ART_MAGENTO_ADVANCEDPRICES "
      oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
      oRecordSet.DoQuery(Query)
      If oRecordSet.RecordCount > 0 Then
        oRecordSet.MoveFirst()
        Dim webRequest As HttpWebRequest
        'rpaya(20220329)
        'Dim strURL As String = URL & "default/async/bulk/V1/products"
        Dim strURL As String = URL & "V1/products/tier-prices"
        'Dim customers As New List(Of Wrapper_Client)

        Dim tierpri As New Wrapper_TierPrice

        Dim ServiceCall As String
        Dim data As Byte()
        Dim stream As Stream

        While Not oRecordSet.EoF

          tierpri = New Wrapper_TierPrice
          Trace.WriteLineIf(ots.TraceInfo, "tierpri.price: " & oRecordSet.Fields.Item(0).Value.ToString.ToString.Replace(",", "."))
          tierpri.price = oRecordSet.Fields.Item(0).Value.ToString.ToString.Replace(",", ".")
          Trace.WriteLineIf(ots.TraceInfo, "tierpri.type: " & oRecordSet.Fields.Item(1).Value.ToString)
          tierpri.price_type = oRecordSet.Fields.Item(1).Value.ToString
          Trace.WriteLineIf(ots.TraceInfo, "tierpri.web id: " & oRecordSet.Fields.Item(2).Value.ToString)
          tierpri.website_id = oRecordSet.Fields.Item(2).Value.ToString
          Trace.WriteLineIf(ots.TraceInfo, "tierpri.sku: " & oRecordSet.Fields.Item(3).Value.ToString)
          tierpri.sku = oRecordSet.Fields.Item(3).Value.ToString
          Trace.WriteLineIf(ots.TraceInfo, "tierpri.group: " & oRecordSet.Fields.Item(4).Value.ToString)
          tierpri.customer_group = oRecordSet.Fields.Item(4).Value.ToString
          Trace.WriteLineIf(ots.TraceInfo, "tierpri.quantity: " & oRecordSet.Fields.Item(5).Value.ToString)
          tierpri.quantity = oRecordSet.Fields.Item(5).Value.ToString


          '**************************************************************************************************************************************
          'Actualizar el precio
          '*************************************************************************************************************************************
          ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf AcceptAllCertifications)
          ServicePointManager.Expect100Continue = True
          ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
          ServiceCall = JsonConvert.SerializeObject(tierpri)
          'ServiceCall = "{ ""customer"": " & ServiceCall.Replace("email"":", "email"":""").Replace(",""first", """,""first") & ",""password"":""" & oRecordSet.Fields.Item("password").Value.ToString & """}"
          ServiceCall = "{ ""prices"": [" & ServiceCall & "] }"

          Trace.WriteLineIf(ots.TraceInfo, "Call to Magento: " & ServiceCall)
          data = Encoding.UTF8.GetBytes(ServiceCall)
          webRequest = HttpWebRequest.Create(strURL)
          webRequest.AllowAutoRedirect = True
          webRequest.Timeout = -1
          webRequest.Method = WebRequestMethods.Http.Post
          webRequest.ContentLength = data.Length
          webRequest.ContentType = "application/json"
          webRequest.Headers.Set("Authorization", "Bearer " & Token)
          stream = webRequest.GetRequestStream()
          stream.Write(data, 0, data.Length)
          stream.Close()
          stream.Dispose()
          Try
            Dim webResponse As HttpWebResponse = webRequest.GetResponse()
            If webResponse.StatusCode = HttpStatusCode.OK Then
              Dim receiveStream As Stream = webResponse.GetResponseStream
              Dim readStream As StreamReader = New StreamReader(receiveStream, Encoding.UTF8)
              Dim respuesta As String = readStream.ReadToEnd
              Dim sCorreo As String = String.Empty

              'Dim Resp = JsonConvert.DeserializeObject(respuesta)
              'Dim Billing_street As String = Resp("billing_address")("street").ToString.Replace("[", "").Replace("]", "")
              'Dim Shipping_street As String = Resp("extension_attributes")("shipping_assignments")(0)("shipping")("address")("street").ToString.Replace("[", "").Replace("]", "")
              Dim ouser As New Wrapper_ClientRes
              ouser = JsonConvert.DeserializeObject(Of Wrapper_ClientRes)(respuesta)
              Trace.WriteLineIf(ots.TraceInfo, "MAGENTO Answer: " & respuesta)
              '"@ART_MAGENTO_PRICES"
              Trace.WriteLineIf(ots.TraceInfo, "Price with Item Code: " & oRecordSet.Fields.Item(3).Value.ToString & " updated to MAGENTO")
              'Dim QueryUpdate As String = "UPDATE OITM SET ""U_ART_mgtBasePrice"" = '" & basepri.price.ToString & "' WHERE ""ItemCode"" = '" & oRecordSet.Fields.Item(1).Value.ToString & "'"
              'Dim oRs As SAPbobsCOM.Recordset
              'oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

              'oRs.DoQuery(QueryUpdate)



              'System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
              'Else
              '  Throw New Exception(webResponse.StatusCode & "-" & webResponse.StatusDescription)
            End If
            'Dim QueryUpdate As String = "UPDATE OCRD SET ""U_ART_MagentoSync"" = 'Y' WHERE ""CardCode"" = '" & oRecordSet.Fields.Item("CardCode").Value.ToString.Trim & "'"



            webResponse.Close()
          Catch ex As Exception
            RECORD_ERRORS(ex.Message, "UPDATE tier prices")
            Trace.WriteLineIf(ots.TraceError, ex.Message)
            Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
          Finally
          End Try
          stream = Nothing

          oRecordSet.MoveNext()
        End While


      Else
        Trace.WriteLineIf(ots.TraceInfo, "There are no tier prices needed to sync with Magento")
      End If
      oResponse = New Helpers.UpdatePricesResponse With {
                              .ExecutionSuccess = True,
                              .FailureReason = String.Empty
                              }
    Catch ex As Exception
      Trace.WriteLineIf(ots.TraceError, ex.Message)
      Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
      oResponse = New Helpers.UpdatePricesResponse With {
                  .ExecutionSuccess = False,
                  .FailureReason = ex.Message
                  }

    End Try
    Return oResponse

  End Function

#End Region


#Region "Funciones y Procedimientos"

  Public Function AcceptAllCertifications(sender As Object, certification As X509Certificate,
                  chain As X509Chain, sslPolicyErrors As SslPolicyErrors) As Boolean
    Return True
  End Function

  Private Sub RECORD_ERRORS(ByVal Message As String, ByVal Procedure As String)
    Dim oRs As SAPbobsCOM.Recordset = Nothing
    Try
      oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
      Dim Query As String
      Select Case oCompany.DbServerType
        Case SAPbobsCOM.BoDataServerTypes.dst_HANADB
          Query = "SELECT IFNULL(MAX(""Code""),0) + 1 FROM ""@ART_ERRORS"""
        Case Else
          Query = "SELECT ISNULL(MAX(Code),0) + 1 FROM [@ART_ERRORS]"
      End Select
      oRs.DoQuery(Query)
      Dim Contador As String = oRs.Fields.Item(0).Value.ToString.Trim
      Dim oUserTable = oCompany.UserTables.Item("ART_ERRORS")
      Dim Fecha As String = Now.ToString("yyyy/MM/dd")
      oUserTable.UserFields.Fields.Item("U_ART_Date").Value = Fecha
      oUserTable.UserFields.Fields.Item("U_ART_Time").Value = Now.ToShortTimeString
      oUserTable.UserFields.Fields.Item("U_ART_ErrorMessage").Value = Message
      oUserTable.UserFields.Fields.Item("U_ART_Procedure").Value = Procedure
      oUserTable.Name = Contador
      If oUserTable.Add <> 0 Then Throw New Exception(oCompany.GetLastErrorDescription)
    Catch ex As Exception
      Trace.WriteLineIf(ots.TraceError, "Error al registrar los errores en la tabla de Errores de SAP: " & ex.Message)
      Trace.WriteLineIf(ots.TraceError, "Error al registrar los errores en la tabla de Errores de SAP: " & ex.StackTrace)
    Finally
      If Not IsNothing(oRs) Then
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
      End If
    End Try
  End Sub

  Private Function UPDATE_ITEMSTATUS(ByVal ListArticulos As Dictionary(Of String, Double)) As Boolean
    Dim result As Boolean = False
    Try
      Dim oRs As SAPbobsCOM.Recordset = Nothing
      Dim fecha As String = Now.ToString("yyyyMMdd")
      For Each element In ListArticulos
        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim Query As String = "UPDATE OITM SET ""U_ART_mgtStockDate"" = '" & fecha & "', ""U_ART_magentoStock"" = " & CInt(element.Value) & " WHERE ""ItemCode"" = '" & element.Key & "'"
        Trace.WriteLineIf(ots.TraceInfo, "Query UPDATE: " & Query)
        oRs.DoQuery(Query)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
        oRs = Nothing
      Next
      result = True
    Catch ex As Exception
      Trace.WriteLineIf(ots.TraceError, ex.Message)
      Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
    End Try
    Return result
  End Function

  Private Function UPDATE_ORDERLINE_STATUS(ByVal ListaDocumentos As Dictionary(Of String, List(Of String))) As Boolean
    'UPDATE RDR1 SET "U_ART_DeliveryDate" = '' WHERE "DocEntry" = Integer AND "U_WebOrderItemId" IN ('String')
    Dim result As Boolean = False
    Try
      Dim oRs As SAPbobsCOM.Recordset = Nothing
      For Each element In ListaDocumentos
        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim ListaLineas As List(Of String) = element.Value
        Dim sListaLineas As String = String.Empty
        For x As Integer = 0 To ListaLineas.Count - 1
          sListaLineas += "'" & ListaLineas.Item(x).ToString.Trim & "'"
          If x < ListaLineas.Count - 1 Then sListaLineas += ","
        Next
        Dim Query As String = "UPDATE RDR1 SET ""U_UpdateDDMagento"" = 'N' WHERE ""DocEntry"" = '" & element.Key & "' AND ""U_WebOrderItemId"" IN (" & sListaLineas & ")"
        Trace.WriteLineIf(ots.TraceInfo, "Query UPDATE: " & Query)
        oRs.DoQuery(Query)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
        oRs = Nothing
      Next
      result = True
    Catch ex As Exception
      Trace.WriteLineIf(ots.TraceError, ex.Message)
      Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
    End Try
    Return result

  End Function

  Private Function ASSIGN_WAREHOUSE_TO_ORDER(ByRef oOrder As SAPbobsCOM.Documents) As Boolean
    'Si el cliente tiene informado el campo U_DflWarehouse en su ficha, asignamos ese almac�n a todas las l�neas y listo.
    'Siempre vamos a priorizar el almac�n 01. Pero, si conseguimos un almac�n que pueda surtir a todas las l�neas del pedido que NO sea el 01, se usa ese.
    'El Almac�n 01 se refiere ahora mismo al almac�n por defecto del Sistema OADM.DfltWhs. 
    Dim Result As Boolean
    Dim oSqlQuery As New DNAUtils.DNASQLUtils.SQLQuery
    Dim oRs As SAPbobsCOM.Recordset = Nothing
    Try
      Dim CardCode As String = oOrder.CardCode
      'Sacar almac�n por defecto de la empresa
      Dim Query As String = "SELECT T0.""DfltWhs"" FROM OADM T0"
      oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
      oRs.DoQuery(Query)
      Dim AlmacenDefectoEmpresa As String = String.Empty
      Dim AlmacenAAsignar As String = String.Empty
      If Not IsNothing(oRs) AndAlso oRs.RecordCount > 0 Then
        AlmacenDefectoEmpresa = oRs.Fields.Item(0).Value.ToString.Trim
        For i As Integer = 0 To oOrder.Lines.Count - 1
          oOrder.Lines.SetCurrentLine(i)
          oOrder.Lines.WarehouseCode = AlmacenDefectoEmpresa
        Next
      Else
        'Sacar el Almac�n del Cliente (si lo tiene)
        Query = "SELECT T0.""U_DflWarehouse"" FROM OCRD T0 WHERE T0.""CardCode"" ='" & CardCode & "'"
        oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRs.DoQuery(Query)
        Dim Asignado As Boolean = False
        If Not IsNothing(oRs) AndAlso oRs.RecordCount > 0 Then
          If Not oRs.Fields.Item(0).Value.ToString <> "" Then
            'Asignamos el Almac�n que tenga el cliente a las l�neas del pedido
            Dim Almacen As String = oRs.Fields.Item(0).Value.ToString.Trim
            Asignado = True
            For i As Integer = 0 To oOrder.Lines.Count - 1
              oOrder.Lines.SetCurrentLine(i)
              oOrder.Lines.WarehouseCode = Almacen
            Next
          End If
        End If
        If Not Asignado Then
          Dim Exists As New List(Of Existencias)
          Dim AsignarAlmacenporDefecto As Boolean = False
          For i As Integer = 0 To oOrder.Lines.Count - 1
            oOrder.Lines.SetCurrentLine(i)
            Query = "SELECT T0.""WhsCode"" FROM OITW T0 INNER JOIN OWHS T1 ON T0.""WhsCode"" = T1.""WhsCode""  WHERE T0.""ItemCode"" = '" & oOrder.Lines.ItemCode & "' AND (T0.""OnHand"" - T0.""IsCommited"") > " & oOrder.Lines.Quantity & " AND T1.""U_AGL_WebStock"" = 'Y' AND T0.""WhsCode""<>'" & AlmacenDefectoEmpresa & "'  ORDER BY(T0.""OnHand"" - T0.""IsCommited"") DESC"
            oRs.DoQuery(Query)
            If oRs.RecordCount > 0 Then
              For x As Integer = 0 To oRs.RecordCount - 1
                Dim Ex As New Existencias
                Ex.ItemCode = oOrder.Lines.ItemCode
                Ex.WhsCode = oRs.Fields.Item(0).Value.ToString.Trim
                Exists.Add(Ex)
              Next
            Else
              AsignarAlmacenporDefecto = True
              'Si el RS viene vac�o, implica que no hay almac�n que suministre a la l�nea. Por tanto usamos el 01 para todas las l�neas
            End If
          Next
          If Exists.Count > 0 AndAlso Not AsignarAlmacenporDefecto Then
            'Revisamos cu�l de los almacenes es gen�rico para todas las l�neas del documento
            'Ordenamos la lista por Almac�n. Si el n�mero de veces que se repite uno coincide con el n�mero de l�neas del documento, pues agarramos ese.
            Exists = Exists.OrderBy(Function(R) R.WhsCode).ToList
            Dim AlmacenAnterior As String = String.Empty
            Dim Contador As Integer = 0
            Dim ConseguiAlmacen As Boolean = False
            For i As Integer = 0 To Exists.Count - 1
              Dim AlmacenActual As Integer = Exists(i).WhsCode
              If AlmacenActual = AlmacenAnterior Or String.IsNullOrEmpty(AlmacenAnterior) Then
                Contador += 1
              Else
                Contador = 1
              End If
              AlmacenAnterior = AlmacenActual
              'Si el contador llega al n�mero de l�neas del documento, entonces asigno ese almac�n.
              If Contador = oOrder.Lines.Count - 1 Then
                ConseguiAlmacen = True
                Exit For
              End If
            Next
            'Si no ha conseguido Almac�n, se asigna el Almac�n por defecto
            If ConseguiAlmacen Then
              AlmacenAAsignar = AlmacenAnterior
            Else
              AlmacenAAsignar = AlmacenDefectoEmpresa
            End If
          Else
            AlmacenAAsignar = AlmacenDefectoEmpresa
          End If
          For i As Integer = 0 To oOrder.Lines.Count - 1
            oOrder.Lines.SetCurrentLine(i)
            oOrder.Lines.WarehouseCode = AlmacenAAsignar
          Next
        End If
      End If
      Result = True
    Catch ex As Exception
      Trace.WriteLineIf(ots.TraceError, ex.Message)
      Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
      Result = False
    Finally
      If Not IsNothing(oRs) Then
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
        oRs = Nothing
      End If
    End Try
    Return Result
  End Function

  Private Function ADD_NEW_CUSTOMER(ByVal oPedido As Wrapper_OrdersCabecera) As String
    Dim result As String = String.Empty
    Try
      Trace.WriteLineIf(ots.TraceInfo, "New customer must be created...")
      Dim oBusinessPartner As SAPbobsCOM.BusinessPartners
      oBusinessPartner = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
      oBusinessPartner.Series = 74 'La serie la indic� Laura por correo
      Trace.WriteLineIf(ots.TraceInfo, "Series = 74")
      oBusinessPartner.CardName = oPedido.customer_firstname & " " & oPedido.customer_lastname
      Trace.WriteLineIf(ots.TraceInfo, oBusinessPartner.CardCode)
      Dim LangCode As Integer = RETURN_LANGUAGE_CODE(oPedido.store_name.Substring(oPedido.store_name.Length - 2, 2))
      If LangCode = -1 Then Throw New Exception("No languague defined on SAP with the Code sent from MAGENTO")
      oBusinessPartner.LanguageCode = LangCode
      Trace.WriteLineIf(ots.TraceInfo, "Languaje Code: " & LangCode)
      oBusinessPartner.GroupCode = 103 'Pasado por correo por Laura
      Trace.WriteLineIf(ots.TraceInfo, "GroupCode = 103")
      oBusinessPartner.EmailAddress = oPedido.customer_email
      Trace.WriteLineIf(ots.TraceInfo, "Email: " & oBusinessPartner.EmailAddress)
      oBusinessPartner.Phone1 = oPedido.billing_address.telephone
      Trace.WriteLineIf(ots.TraceInfo, "Phone: " & oBusinessPartner.Phone1)
      oBusinessPartner.UserFields.Fields.Item("U_WEBSITE").Value = "B2C"
      Trace.WriteLineIf(ots.TraceInfo, "U_WEBSITE = B2C")

      '220715 
      oBusinessPartner.UserFields.Fields.Item("U_WBCUSTID").Value = oPedido.customer_id
      Trace.WriteLineIf(ots.TraceInfo, "U_WBCUSTID = " & oBusinessPartner.UserFields.Fields.Item("U_WBCUSTID").Value)
      '220914
      oBusinessPartner.PayTermsGrpCode = 30
      Trace.WriteLineIf(ots.TraceInfo, "GroupNum = 30")
      oBusinessPartner.Industry = 74
      Trace.WriteLineIf(ots.TraceInfo, "Industry = 74")

      '220715  persona de contacto
      oBusinessPartner.ContactEmployees.Title = "Web"
      oBusinessPartner.ContactEmployees.FirstName = oPedido.customer_firstname
      oBusinessPartner.ContactEmployees.LastName = oPedido.customer_lastname
      oBusinessPartner.ContactEmployees.E_Mail = oPedido.customer_email
      oBusinessPartner.ContactEmployees.Add()
      'oBusinessPartner.ContactPerson = "Web Customer"
      'Trace.WriteLineIf(ots.TraceInfo, "Web Customer = " & oBusinessPartner.ContactPerson)

      Dim Company As String = String.Empty
      If Not IsNothing(oPedido.billing_address.company) AndAlso Not String.IsNullOrEmpty(oPedido.billing_address.company) Then
        Company = oPedido.billing_address.company
        Trace.WriteLineIf(ots.TraceInfo, "Company: " & Company)
      End If
      If Company.Length > 50 Then Company = Company.Substring(0, 49)
      Dim Nombre As String = oPedido.customer_firstname & " " & oPedido.customer_lastname
      If Nombre.Length > 50 Then Nombre = Nombre.Substring(0, 49)

      'Direcci�n de facturaci�n
      oBusinessPartner.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo
      oBusinessPartner.Addresses.City = oPedido.billing_address.city
      Trace.WriteLineIf(ots.TraceInfo, "Billing Address - City: " & oBusinessPartner.Addresses.City)
      oBusinessPartner.Addresses.ZipCode = oPedido.billing_address.postcode
      Trace.WriteLineIf(ots.TraceInfo, "Billing Address - Postal Code: " & oBusinessPartner.Addresses.ZipCode)
      oBusinessPartner.Addresses.County = oPedido.billing_address.region_code
      Trace.WriteLineIf(ots.TraceInfo, "Billing Address - County: " & oBusinessPartner.Addresses.County)
      oBusinessPartner.Addresses.AddressName = "Billing"
      oBusinessPartner.Addresses.Block = oPedido.billing_address.telephone
      Trace.WriteLineIf(ots.TraceInfo, "Billing Address - Block: " & oBusinessPartner.Addresses.Block)
      oBusinessPartner.Addresses.UserFields.Fields.Item("U_TelNo").Value = oPedido.billing_address.telephone
      Trace.WriteLineIf(ots.TraceInfo, "Billing Address - U_TelNo: " & oBusinessPartner.Addresses.UserFields.Fields.Item("U_TelNo").Value)
      oBusinessPartner.Addresses.AddressName2 = Company
      Trace.WriteLineIf(ots.TraceInfo, "Billing Address - AddressName2: " & oBusinessPartner.Addresses.AddressName2)
      oBusinessPartner.Addresses.AddressName3 = Nombre
      Trace.WriteLineIf(ots.TraceInfo, "Billing Address - AddresName3: " & oBusinessPartner.Addresses.AddressName3)
      Trace.WriteLineIf(ots.TraceInfo, "Billing Address - WBCUSTADDID: " & oPedido.billing_address.customer_address_id)
      oBusinessPartner.Addresses.UserFields.Fields.Item("U_WBCUSTADDID").Value = oPedido.billing_address.customer_address_id
      'Trace.WriteLineIf(ots.TraceInfo, "Billing Address - WBCUSTADDID: " & oBusinessPartner.Addresses.UserFields.Fields.Item("U_WBCUSTADDID").Value)
      Dim DirFact As String = String.Empty
      For Each part In oPedido.billing_address.street
        DirFact += part & " "
      Next
      Trace.WriteLineIf(ots.TraceInfo, "DirFact: " & DirFact)
      If DirFact.Length > 100 Then DirFact = DirFact.Substring(0, 99)
      oBusinessPartner.Addresses.Street = DirFact
      Trace.WriteLineIf(ots.TraceInfo, "Billing Address - Street: " & oBusinessPartner.Addresses.Street)

      oBusinessPartner.Addresses.Add()

      'Direcci�n de Entrega



      oBusinessPartner.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_ShipTo
      oBusinessPartner.Addresses.City = oPedido.extension_attributes.shipping_assignments(0).shipping.address.city
      Trace.WriteLineIf(ots.TraceInfo, "Shipping Address - City: " & oBusinessPartner.Addresses.City)
      oBusinessPartner.Addresses.ZipCode = oPedido.extension_attributes.shipping_assignments(0).shipping.address.postcode
      Trace.WriteLineIf(ots.TraceInfo, "Shipping Address - Postal Code: " & oBusinessPartner.Addresses.ZipCode)
      oBusinessPartner.Addresses.County = oPedido.extension_attributes.shipping_assignments(0).shipping.address.region_code
      Trace.WriteLineIf(ots.TraceInfo, "Shipping Address - County: " & oBusinessPartner.Addresses.County)
      oBusinessPartner.Addresses.AddressName = "Shipping"
      oBusinessPartner.Addresses.Block = oPedido.extension_attributes.shipping_assignments(0).shipping.address.telephone
      Trace.WriteLineIf(ots.TraceInfo, "Shipping Address - Block: " & oBusinessPartner.Addresses.Block)
      oBusinessPartner.Addresses.UserFields.Fields.Item("U_TelNo").Value = oPedido.extension_attributes.shipping_assignments(0).shipping.address.telephone
      Trace.WriteLineIf(ots.TraceInfo, "Shipping Address - U_TelNo: " & oBusinessPartner.Addresses.UserFields.Fields.Item("U_TelNo").Value)
      oBusinessPartner.Addresses.AddressName2 = Company
      Trace.WriteLineIf(ots.TraceInfo, "Shipping Address - AddressName2: " & oBusinessPartner.Addresses.AddressName2)
      oBusinessPartner.Addresses.AddressName3 = Nombre
      Trace.WriteLineIf(ots.TraceInfo, "Shipping Address - AddresName3: " & oBusinessPartner.Addresses.AddressName3)
      oBusinessPartner.Addresses.UserFields.Fields.Item("U_WBCUSTADDID").Value = CStr(oPedido.extension_attributes.shipping_assignments(0).shipping.address.customer_address_id)
      Trace.WriteLineIf(ots.TraceInfo, "Shipping Address - WBCUSTADDID: " & oBusinessPartner.Addresses.UserFields.Fields.Item("U_WBCUSTADDID").Value)
      Dim DirEnv As String = String.Empty
      For Each part In oPedido.extension_attributes.shipping_assignments(0).shipping.address.street
        DirEnv += part & " "
      Next
      Trace.WriteLineIf(ots.TraceInfo, "DirEnv: " & DirEnv)
      If DirEnv.Length > 100 Then DirEnv = DirEnv.Substring(0, 99)
      oBusinessPartner.Addresses.Street = DirEnv
      Trace.WriteLineIf(ots.TraceInfo, "Shipping Address - Street: " & oBusinessPartner.Addresses.Street)

      If oBusinessPartner.Add <> 0 Then Throw New Exception("Error adding new Customer in SAP: " & oCompany.GetLastErrorDescription)
      result = oCompany.GetNewObjectKey
      Trace.WriteLineIf(ots.TraceInfo, "Customer added with Code: " & result)
    Catch ex As Exception
      Trace.WriteLineIf(ots.TraceError, ex.Message)
      Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
      result = "-1"
    End Try
    Return result
  End Function

  Private Function RETURN_CONTACTPERSON(ByVal CardCode As String, ByVal Email As String) As Integer
    Dim sQuery As String = "SELECT ""CntctCode"" FROM OCPR WHERE ""CardCode"" ='" & CardCode & "' AND  ""E_MailL"" ='" & Email & "'"
    Dim oRs As SAPbobsCOM.Recordset
    oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    Dim result As Integer = -1
    Try
      oRs.DoQuery(sQuery)
      If oRs.RecordCount > 0 Then
        result = CInt(oRs.Fields.Item(0).Value.ToString.Trim)
      End If
    Catch ex As Exception
      Trace.WriteLineIf(ots.TraceError, ex.Message)
      Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
    Finally
      If Not IsNothing(oRs) Then
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
        oRs = Nothing
      End If
    End Try
    Return result
  End Function

  Private Function RETURN_LANGUAGE_CODE(ByVal LangShotName) As Integer
    Dim result As Integer = -1
    Dim oRs As SAPbobsCOM.Recordset = Nothing
    Try
      oRs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
      Dim Query As String = "SELECT T0.""Code"" FROM OLNG T0 WHERE T0.""ShortName"" ='" & LangShotName & "'"
      oRs.DoQuery(Query)
      If oRs.RecordCount > 0 Then
        result = CInt(oRs.Fields.Item(0).Value.ToString.Trim)
      Else
        result = -1
      End If
    Catch ex As Exception
      Trace.WriteLineIf(ots.TraceError, ex.Message)
      Trace.WriteLineIf(ots.TraceError, ex.StackTrace)
      result = -1
    Finally
      If Not IsNothing(oRs) Then
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
        oRs = Nothing
      End If
    End Try
    Return result

  End Function

#End Region


End Class