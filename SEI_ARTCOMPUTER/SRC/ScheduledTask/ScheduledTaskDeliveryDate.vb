Imports ModuleBase
Imports SAPbobsCOM

Public Class ScheduledTaskDeliveryDate

    Inherits ModuleBase.AbstractScheduledTask

    Public Overrides Function Execute(oCompany As Company, params As Hashtable) As ScheduledTaskResult
        Dim oResponse As New Helpers.UpdateDeliveryDateCABResponse
        oResponse = SEI_ARTCOMPUTER_Services.GetInstance(oCompany).UPDATEDeliveryDate_NOSYNC()
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
        Return New IScheduledTask.JobInfo() With {.TaskName = "SincDeliveryDate"}
    End Function
End Class
