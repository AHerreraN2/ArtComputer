'----------------------------------------------------------------------------------------------------
'This is an automagically generated template for an Addonia AddOn Module.
'Please refer to INSTRUCTIONS.txt for further help.
'
'ModuleName: SEI_ARTCOMPUTER
'Author: jsalgueiro.SEIDORBCN  @  PC03D36U
'Date: 1/10/2022 7:36:27 PM
'----------------------------------------------------------------------------------------------------
Imports System.IO
Imports System.Net
Imports System.Text
Imports ModuleBase
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SAPbouiCOM

Public Class SEI_ARTCOMPUTER
  Inherits ModuleBase.AbstractModuleModel

#Region "Implementation Module Property"
  Private _implementation As SEI_ARTCOMPUTERImpl
  Public Property oImplementation() As SEI_ARTCOMPUTERImpl
    Get
      Return _implementation
    End Get
    Set(ByVal value As SEI_ARTCOMPUTERImpl)
      _implementation = value
    End Set
  End Property

  Public Overrides Sub InitializeImplementation()
    _implementation = New SEI_ARTCOMPUTERImpl
    _implementation.initializeModule(oApplication, oCompany, oAddOnService)
    _implementation.oModule = Me
  End Sub
#End Region


  Public Overrides Sub queryMetadataCreationConfig(ByRef oMetadataCreationSelector As ModuleBase.cMetadaCreationSelector)
    oMetadataCreationSelector.ShouldCreateMetadata = True
    oMetadataCreationSelector.ShouldCreateUserTables = True
    oMetadataCreationSelector.ShouldCreateUserFields = True
  End Sub

  Public Overrides Function getFilterCollection() As List(Of cEventFilter)
    Return MyBase.getFilterCollection()
  End Function

  'TODO: Write you module here...




  Public Overrides Sub oApplication_MenuEvent(ByRef pVal As MenuEvent, ByRef BubbleEvent As Boolean)
    MyBase.oApplication_MenuEvent(pVal, BubbleEvent)
        Dim oUser As SAPbobsCOM.Users
        oUser = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers)
        oUser.GetByKey(1)


        'If pVal.MenuUID = "257" And pVal.BeforeAction Then
        '        'Cargo el borrador del Albarán.
        '        Dim oDraft As SAPbobsCOM.Documents
        '        oDraft = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
        '        oDraft.GetByKey(2)
        '        'Cargo la Factura de Reserva para sacar los datos.
        '        Dim oFactura As SAPbobsCOM.Documents
        '        oFactura = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
        '        oFactura.GetByKey(367)
        '        'Borro las líneas del borrador
        '        For i As Integer = oDraft.Lines.Count - 1 To 0 Step -1
        '            oDraft.Lines.SetCurrentLine(i)
        '            oDraft.Lines.Delete()
        '        Next
        '        For j As Integer = 0 To oFactura.Lines.Count - 1
        '            oFactura.Lines.SetCurrentLine(j)
        '            If Not String.IsNullOrEmpty(oDraft.Lines.ItemCode) Then oDraft.Lines.Add()
        '            oDraft.Lines.BaseEntry = oFactura.DocEntry
        '            oDraft.Lines.BaseLine = oFactura.Lines.LineNum
        '            oDraft.Lines.BaseType = 13
        '        Next
        '        oDraft.Update()
        '        Try
        '            If oDraft.SaveDraftToDocument() <> 0 Then
        '                oApplication.MessageBox("Error: " & oCompany.GetLastErrorDescription)
        '            End If
        '        Catch ex As Exception
        '        End Try




        '    End If
    End Sub



End Class
