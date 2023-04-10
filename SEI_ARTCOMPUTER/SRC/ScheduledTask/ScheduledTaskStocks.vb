Imports ModuleBase
Imports SAPbobsCOM

Public Class ScheduledTaskStocks
    Inherits ModuleBase.AbstractScheduledTask

    Public Overrides Function Execute(oCompany As Company, params As Hashtable) As ScheduledTaskResult
        Dim oResponse As New Helpers.UpdateStocksResponse
        oResponse = SEI_ARTCOMPUTER_Services.GetInstance(oCompany).UPDATEStocks_NOSYNC()
        Dim result As ScheduledTaskResult = New ScheduledTaskResult
        If oResponse.ExecutionSuccess Then
            result.ExecutionSuccess = True
        Else
            result.ExecutionSuccess = False
            result.ExecutionSuccess = oResponse.FailureReason
        End If
        Return result

    End Function

    Public Overrides Function GetJobInfo() As IScheduledTask.JobInfo
        Return New IScheduledTask.JobInfo() With {.TaskName = "SincStocks"}
    End Function
End Class
