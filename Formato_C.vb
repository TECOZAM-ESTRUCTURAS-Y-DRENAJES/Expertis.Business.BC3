Public Class Formato_C
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub
    Private Const cnEntidad As String = "FORMATO_C"

#Region " RegisterUpdateTask "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarID)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarCodigoInicial)
    End Sub

    <Task()> Public Shared Sub AsignarID(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("ID")) = 0 OrElse IsDBNull(data("ID")) Then
                data("ID") = AdminData.GetAutoNumeric
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarCodigoInicial(ByVal data As DataRow, ByVal services As ServiceProvider)
        'If data.RowState = DataRowState.Added Then
        '    data("CODIGOINICIAL") = data("CODIGO")
        'End If
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
