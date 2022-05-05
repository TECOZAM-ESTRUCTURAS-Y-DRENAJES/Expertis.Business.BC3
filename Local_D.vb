Public Class Local_D
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Private Const cnEntidad As String = "LOCAL_D"
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

#Region " RegisterUpdateTask "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarID)
    End Sub

    <Task()> Public Shared Sub AsignarID(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("ID")) = 0 OrElse IsDBNull(data("ID")) Then
                data("ID") = AdminData.GetAutoNumeric
            End If
        End If
    End Sub

#End Region

#Region " RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarID, data, services)
    End Sub

#End Region

End Class
