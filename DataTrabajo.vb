Public Class DataTrabajo

    Public NumPresup As String
    Public IDTipoObra As String
    Public IDTipoTrabajo As String
    Public IDSubTipoTrabajo As String
    Public IDSubSubTipoTrabajo As String
    Public dblAcumular As Boolean
    Public nivel As Integer
    Public orden As Integer
    Public Rama As Integer
    Public IDPadre As Integer
    Public IDHijo As Integer
    Public IDFiltro As Integer
    Public CodigoPadre As String
    Public CodigoHijo As String
    Public strIDArticulo As String
    Public strTipoHora As String
    Public dtTrabajosN As DataTable
    Public dtObraTipoTrabajoN As DataTable
    Public dtObraSubTipoTrabajoN As DataTable
    Public dtObraSubSubTipoTrabajoN As DataTable
    Public dtMaterialesN As DataTable
    Public dtManoObraN As DataTable
    Public dtCentroTrabajoN As DataTable
    Public dtVariosTrabajoN As DataTable
    Public dtMedicionesN As DataTable

    Public blnImportarMateriales As Boolean
    Public blnImportarMOD As Boolean
    Public blnImportarCentros As Boolean
    Public blnImportarMediciones As Boolean
    Public cantidad As Double


End Class
