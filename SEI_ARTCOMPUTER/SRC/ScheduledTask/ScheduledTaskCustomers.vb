Imports ModuleBase
Imports SAPbobsCOM
Public Class ScheduledTaskCustomers
  Inherits ModuleBase.AbstractScheduledTask

  Public Overrides Function Execute(oCompany As Company, params As Hashtable) As ScheduledTaskResult
    Dim oResponse As New Helpers.UpdateClientsResponse
    oResponse = SEI_ARTCOMPUTER_Services.GetInstance(oCompany).UPDATEclients_NOSYNC()
    Dim result As ScheduledTaskResult = New ScheduledTaskResult
    If oResponse.ExecutionSuccess Then
      result.ExecutionSuccess = True
    Else
      result.ExecutionSuccess = False
      result.FailureReason = oResponse.FailureReason
    End If
    Return result
  End Function

  Public Overrides Function GetJobInfo() As IScheduledTask.JobInfo
    Return New IScheduledTask.JobInfo() With {.TaskName = "UpdateCustomers"}
  End Function
End Class
