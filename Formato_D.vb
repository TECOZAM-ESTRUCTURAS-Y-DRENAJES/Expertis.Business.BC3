Public Class Formato_D
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub
    Private Const cnEntidad As String = "FORMATO_D"

#Region " RegisterUpdateTask "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarID)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarCodigoPadreInicial)
    End Sub

    <Task()> Public Shared Sub AsignarID(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("ID")) = 0 OrElse IsDBNull(data("ID")) Then
                data("ID") = AdminData.GetAutoNumeric
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarCodigoPadreInicial(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            data("CODIGOPADREINICIAL") = data("CODIGOPADRE")
        End If
    End Sub

#End Region

#Region " AddNewForm"

    Public Overrides Function AddNewForm() As DataTable
        Dim dt As DataTable = MyBase.AddNewForm
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf FillDefaultValues, dt.Rows(0), New ServiceProvider)
        Return dt
    End Function

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarID, data, services)
    End Sub

#End Region

End Class
