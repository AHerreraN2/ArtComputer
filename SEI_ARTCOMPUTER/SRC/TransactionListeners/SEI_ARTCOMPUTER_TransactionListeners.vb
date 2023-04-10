Public Class SEI_ARTCOMPUTER_TransactionListeners
    Inherits ModuleBase.AbstractTransactionToDIListenerModel

    Protected Overrides Function getListeningObjectTypes() As System.Collections.Generic.List(Of String)
        'Add listening objects here (AddListeningObject)
        Return MyBase.getListeningObjectTypes()
    End Function

    Public Overrides Function ListenPostTransaction(ByVal oCompany As SAPbobsCOM.Company, ByVal dbName As String, ByVal object_type As String, ByVal transaction_type As String, ByVal num_of_cols_in_key As Integer, ByVal list_of_key_cols_tab_del As String, ByVal list_of_cols_val_tab_del As String) As Boolean
        Return True
    End Function

    Public Overrides Function ListenTransaction(ByVal oCompany As SAPbobsCOM.Company, ByVal dbName As String, ByVal object_type As String, ByVal transaction_type As String, ByVal num_of_cols_in_key As Integer, ByVal list_of_key_cols_tab_del As String, ByVal list_of_cols_val_tab_del As String) As String
        Return ""
    End Function

End Class
