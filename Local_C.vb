Public Class Local_C
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub
    Private Const cnEntidad As String = "LOCAL_C"

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

    Public Sub EJECUTARFORMULAS()
        Dim Med As Formato_M
        Dim dtt As DataTable
        Dim Formula As New cFormulas

        Med = New Formato_M
        dtt = Med.Filter(New NumberFilterItem("Tipo", 3))


        For Each row As DataRow In dtt.Rows
            row.Item("Medicion") = xRound(Nz(Formula.Calcular(row.Item("Formula")), 0), 3)
        Next
        Med.Update(dtt)

    End Sub

    Public Sub LecturaFichero(ByVal ArrayTexto As String(), ByVal NumPresup As Integer, _
                              Optional ByVal blnImportarMateriales As Boolean = True, _
                              Optional ByVal blnCrearMateriales As Boolean = True, _
                              Optional ByVal blnImportarMOD As Boolean = True, _
                              Optional ByVal blnImportarCentros As Boolean = True, _
                              Optional ByVal blnImportarMediciones As Boolean = True)

        Dim texto As String
        Dim NewRow As DataRow

        AdminData.ExecuteNonQuery("spBorradoBC3", False)

        Dim C As New Local_C
        Dim D As New Local_D
        Dim T As New Local_T
        Dim M As New Local_M

        Dim dttC As DataTable = C.AddNew
        Dim dttD As DataTable = D.AddNew
        Dim dttT As DataTable = T.AddNew
        Dim dttM As DataTable = M.AddNew

        If ArrayTexto.Length > 0 Then
            For Each texto In ArrayTexto
                Select Case Mid(texto, 1, 2)
                    Case "~C"
                        NewRow = dttC.NewRow
                        NewRow.Item("descripcion") = Replace(texto, "'", ",")
                        dttC.Rows.Add(NewRow)
                    Case "~D"
                        NewRow = dttD.NewRow
                        NewRow.Item("descripcion") = Replace(texto, "'", ",")
                        dttD.Rows.Add(NewRow)
                    Case "~T"
                        NewRow = dttT.NewRow
                        NewRow.Item("descripcion") = Replace(texto, "'", ",")
                        dttT.Rows.Add(NewRow)
                    Case "~M"
                        NewRow = dttM.NewRow
                        NewRow.Item("descripcion") = Replace(texto, "'", ",")
                        dttM.Rows.Add(NewRow)
                End Select
            Next

            C.Update(dttC)
            D.Update(dttD)
            T.Update(dttT)
            M.Update(dttM)

            'Desglosamos el Fichero....

            FORMATO_C()
            FORMATO_D()
            FORMATO_M()
            FORMATO_T()

            AdminData.ExecuteNonQuery("spFormulas", False)

            EJECUTARFORMULAS()

            CrearPresupuesto(NumPresup, blnImportarMateriales, blnCrearMateriales, blnImportarMOD, blnImportarCentros, blnImportarMediciones)
        End If
    End Sub
    Public Sub FORMATO_C()
        '~C |{CODIGO\}|UNIDAD|RESUMEN|{PRECIO\}|{FECHA\}|TIPO|CURVA_PRECIO_SIMPLE|CURVA_PRECIO_COMPUESTO
        'Damos el formato a los datos de tipo Concepto...
        Dim barra As Integer
        Dim txtSQL As String
        Dim unidad As String
        Dim Resumen As String
        Dim precio As String
        Dim fecha As String
        Dim tipo As String
        Dim FC As Formato_C
        'Dim LC As Local_C
        Dim dtt As DataTable
        Dim dttIns As DataTable
        Dim texto As String
        Dim codigo As String
        Dim Tamaño As Integer
        Dim TextoInicial As String
        Dim NewRow As DataRow


        Try

            barra = 0

            FC = New Formato_C
            'LC = New Local_C
            dtt = Me.Filter
            dttIns = FC.AddNew

            For Each row As DataRow In dtt.Rows
                codigo = ""
                unidad = ""
                precio = ""
                tipo = ""
                fecha = ""
                Resumen = ""

          
                Tamaño = Len(row.Item("Descripcion"))
                texto = Mid(row.Item("Descripcion"), 4, Tamaño)
                TextoInicial = texto
                '       While texto <> "" And texto <> "||"
                barra = InStr(texto, "|")
                If codigo = "" Then
                    codigo = Replace(Replace(Mid(texto, 1, barra - 1), "\", ""), "#", "@")
                    texto = Mid(texto, barra + 1, Tamaño)
                    barra = InStr(texto, "|")
                End If
                If unidad = "" Then
                    unidad = Mid(texto, 1, barra - 1)
                    If unidad = "" Then
                        unidad = "00"
                    End If
                    texto = Mid(texto, barra + 1, Tamaño)
                    barra = InStr(texto, "|")
                End If
                If Resumen = "" Then
                    If barra > 0 Then
                        Resumen = Mid(texto, 1, barra - 1) '1
                        texto = Mid(texto, barra + 1, Tamaño)
                    Else
                        texto = ""
                    End If
                    If Resumen = "" Then
                        Resumen = " "
                    End If
                    barra = InStr(texto, "|")
                End If
                If precio = "" Then
                    If barra > 0 Then
                        precio = Replace(Mid(texto, 1, barra - 1), ".", ",")
                        texto = Mid(texto, barra + 1, Tamaño)
                    Else
                        texto = ""
                    End If

                    If precio = "" Then
                        precio = 0
                    End If
                    barra = InStr(texto, "|")
                End If
                If fecha = "" Then
                    If barra > 0 Then
                        fecha = Mid(texto, 1, barra - 1)
                        texto = Mid(texto, barra + 1, Tamaño)
                    Else
                        texto = ""
                    End If
                    If fecha = "" Then
                        fecha = Date.Today
                    End If
                    barra = InStr(texto, "|")
                End If

                'FALTA RECORRER LAS CURVAS
                If tipo = "" Then
                    If barra > 0 Then
                        tipo = Mid(texto, 1, barra - 1)
                        texto = Mid(texto, barra + 1, Tamaño)
                    Else
                        texto = ""
                    End If
                    NewRow = dttIns.NewRow
                    NewRow.Item("codigo") = codigo
                    If InStr(codigo, "%") Then
                        NewRow.Item("porcentaje") = 1
                    End If
                    NewRow.Item("unidad") = unidad
                    NewRow.Item("Resumen") = Resumen
                    If InStr(precio, "\") Then
                        precio = Mid(precio, 1, InStr(precio, "\") - 1)
                    End If
                    If Not IsNumeric(precio) Then
                        NewRow.Item("Precio") = 0
                    Else
                        NewRow.Item("Precio") = CDbl(precio)
                    End If

                    If fecha <> "" Then
                        If InStr(fecha, "\") Then
                            fecha = Mid(fecha, 1, InStr(fecha, "\") - 1)
                        End If


                        NewRow.Item("fecha") = fecha
                    End If
                    NewRow.Item("tipo") = tipo
                    NewRow.Item("ID") = row.Item("Id")
                    NewRow.Item("Capitulo") = 1
                    If InStr(TextoInicial, "#", CompareMethod.Text) Then
                        NewRow.Item("Nivel") = 0
                        NewRow.Item("Capitulo") = 1
                    Else
                        NewRow.Item("Capitulo") = 0
                    End If
                    NewRow.Item("codigoinicial") = NewRow.Item("codigo")
                    dttIns.Rows.Add(NewRow)
                End If
            Next

            FC.Update(dttIns)

            FC = Nothing
            dttIns = Nothing
            dtt = Nothing

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub FORMATO_C_old()
        '~C |{CODIGO\}|UNIDAD|RESUMEN|{PRECIO\}|{FECHA\}|TIPO|CURVA_PRECIO_SIMPLE|CURVA_PRECIO_COMPUESTO
        'Damos el formato a los datos de tipo Concepto...
        Dim barra As Integer
        Dim unidad As String
        Dim Resumen As String
        Dim precio As String
        Dim fecha As String
        Dim tipo As String
        Dim FC As Formato_C
        'Dim LC As Local_C
        Dim dtt As DataTable
        Dim dttIns As DataTable
        Dim texto As String
        Dim codigo As String
        Dim Tamaño As Integer
        Dim TextoInicial As String
        Dim NewRow As DataRow

        Try

            barra = 0

            FC = New Formato_C
            'LC = New Local_C
            dtt = Me.Filter
            dttIns = FC.AddNew

            For Each row As DataRow In dtt.Rows
                codigo = ""
                unidad = ""
                precio = ""
                tipo = ""
                fecha = ""
                Resumen = ""

                Tamaño = Len(row.Item("Descripcion"))
                texto = Mid(row.Item("Descripcion"), 4, Tamaño)
                TextoInicial = texto
                While texto <> "" And texto <> "||"
                    barra = InStr(texto, "|")
                    If codigo = "" Then
                        codigo = Replace(Replace(Mid(texto, 1, barra - 1), "\", ""), "#", "@")
                        texto = Mid(texto, barra + 1, Tamaño)
                    Else
                        If unidad = "" Then
                            unidad = Mid(texto, 1, barra - 1)
                            If unidad = "" Then
                                unidad = "00"
                            End If
                            texto = Mid(texto, barra + 1, Tamaño)
                        Else
                            If Resumen = "" Then
                                If barra > 0 Then
                                    Resumen = Mid(texto, 1, barra - 1) '1
                                    texto = Mid(texto, barra + 1, Tamaño)
                                Else
                                    texto = ""
                                End If
                                If Resumen = "" Then
                                    Resumen = " "
                                End If

                            Else
                                If precio = "" Then
                                    If barra > 0 Then
                                        precio = Replace(Mid(texto, 1, barra - 1), ".", ",")
                                        texto = Mid(texto, barra + 1, Tamaño)
                                    Else
                                        texto = ""
                                    End If

                                    If precio = "" Then
                                        precio = 0
                                    End If
                                Else
                                    If fecha = "" Then
                                        If barra > 0 Then
                                            fecha = Mid(texto, 1, barra - 1)
                                            texto = Mid(texto, barra + 1, Tamaño)
                                        Else
                                            texto = ""
                                        End If
                                        If fecha = "" Then
                                            fecha = Date.Today
                                        End If
                                    Else

                                        'FALTA RECORRER LAS CURVAS
                                        If tipo = "" Then
                                            If barra > 0 Then
                                                tipo = Mid(texto, 1, barra - 1)
                                                texto = Mid(texto, barra + 1, Tamaño)
                                            Else
                                                texto = ""
                                            End If
                                            NewRow = dttIns.NewRow
                                            NewRow.Item("codigo") = codigo
                                            If InStr(codigo, "%") Then
                                                NewRow.Item("porcentaje") = 1
                                            End If
                                            NewRow.Item("unidad") = unidad
                                            NewRow.Item("Resumen") = Resumen
                                            If InStr(precio, "\") Then
                                                precio = Mid(precio, 1, InStr(precio, "\") - 1)
                                            End If
                                            If Not IsNumeric(precio) Then
                                                NewRow.Item("Precio") = 0
                                            Else
                                                NewRow.Item("Precio") = CDbl(precio)
                                            End If

                                            If fecha <> "" Then
                                                NewRow.Item("fecha") = fecha
                                            End If
                                            NewRow.Item("tipo") = tipo
                                            NewRow.Item("ID2") = row.Item("Id")
                                            NewRow.Item("Capitulo") = 1
                                            If InStr(TextoInicial, "#", CompareMethod.Text) Then
                                                NewRow.Item("Nivel") = 0
                                                NewRow.Item("Capitulo") = 1
                                            Else
                                                NewRow.Item("Capitulo") = 0
                                            End If
                                            dttIns.Rows.Add(NewRow)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End While
            Next

            FC.Update(dttIns)

            FC = Nothing
            dttIns = Nothing
            dtt = Nothing

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Sub FORMATO_D()
        'Damos el formato a los datos de tipo Desglose...
        Dim barra As Integer
        Dim Parte1 As String
        Dim Parte2 As String
        Dim Parte3 As String
        Dim dtt As DataTable
        Dim dttIns As DataTable
        Dim LD As Local_D
        Dim FD As FORMATO_D
        Dim texto As String
        Dim Insertado As Boolean
        Dim Tamaño As Integer
        Dim TextoInicial As String
        Dim Padre As String
        Dim Orden As Long
        Dim NewRow As DataRow

        Try
            Orden = 10
            barra = 0

            LD = New Local_D
            FD = New Formato_D


            dtt = LD.Filter
            dttIns = FD.AddNew

            For Each row As DataRow In dtt.Rows
                'Primero sacamos el codigo del padre...

                Tamaño = Len(row.Item("Descripcion"))
                texto = Mid(row.Item("Descripcion"), 4, Tamaño)
                TextoInicial = texto

                barra = InStr(texto, "|")
                If barra <> 0 Then
                    Padre = Mid(texto, 1, barra - 1)
                    texto = Mid(texto, barra + 1, Tamaño)

                    'Ahora sacamos los hijos...
                    Parte1 = ""
                    Parte2 = ""
                    Parte3 = ""
                    Insertado = False
                    While texto <> "|" And texto <> "" 'And barra <> 0
                        barra = InStr(texto, "\")
                        If barra = 0 Then barra = 1
                        If Parte1 = "" Then
                            Parte1 = Mid(texto, 1, barra - 1)
                            texto = Mid(texto, barra + 1, Tamaño)
                        Else
                            If Parte2 = "" Then
                                Parte2 = Mid(texto, 1, barra - 1)
                                texto = Mid(texto, barra + 1, Tamaño)
                                If Parte2 = "" Then
                                    Parte2 = "0"
                                End If
                            Else
                                If Parte3 = "" Then
                                    Parte3 = Mid(texto, 1, barra - 1)
                                    texto = Mid(texto, barra + 1, Tamaño)
                                    If Parte3 = "" Then
                                        Parte3 = "0"
                                    End If
                                Else 'Insertamos
                                    NewRow = dttIns.NewRow
                                    NewRow.Item("CodigoPadre") = Replace(Padre, "#", "@")
                                    NewRow.Item("CodigoHijo") = Parte1
                                    NewRow.Item("Factor") = Nz(Trim(Parte2), 0)
                                    If Left(Parte3, 1) = "." Then
                                        Parte3 = "0" & Parte3
                                    End If
                                    NewRow.Item("RENDIMIENTO") = Nz(Trim(Parte3), 0)
                                    dttIns.Rows.Add(NewRow)

                                    Insertado = True
                                    Parte1 = ""
                                    Parte2 = ""
                                    Parte3 = ""
                                    Insertado = False

                                End If
                            End If
                        End If
                        If Not Insertado And Parte1 <> "" And Parte2 <> "" And Parte3 <> "" Then
                            NewRow = dttIns.NewRow
                            NewRow.Item("CodigoPadre") = Replace(Padre, "#", "@")
                            ' NewRow.Item("CodigoHijo") = Parte1
                            If Len(Parte1) > 0 AndAlso InStr(Parte1, "#") Then
                                NewRow.Item("CodigoHijo") = Replace(Parte1, "#", "")
                            Else
                                NewRow.Item("CodigoHijo") = Parte1
                            End If
                            NewRow.Item("Factor") = Nz(Trim(Parte2), 0)
                            If Left(Parte3, 1) = "." Then
                                Parte3 = "0" & Parte3
                            End If

                            NewRow.Item("RENDIMIENTO") = Nz(Replace(Parte3, ".", ","), 0)
                            '   NewRow.Item("Orden") = row.Item("Id")
                            dttIns.Rows.Add(NewRow)
                            Parte1 = ""
                            Parte2 = ""
                            Parte3 = ""

                        End If

                    End While
                End If
                Orden = 10
            Next



            FD.Update(dttIns)
            FD = Nothing
            LD = Nothing
            dtt = Nothing
            dttIns = Nothing
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Sub FORMATO_M()

        'Damos el formato a los datos de tipo Concepto...
        Dim barra As Integer
        Dim CODIGOPADRE As String
        Dim CodigoHijo As String
        Dim Medicion As String
        Dim dtt As DataTable
        Dim dttIns As DataTable
        Dim texto As String
        Dim descripcion As String
        Dim Altura As Double
        Dim Anchura As Double
        Dim Cantidad As Double
        Dim Largo As Double
        Dim tipo As String
        Dim LM As Local_M
        Dim FM As FORMATO_M
        Dim Tamaño As Integer
        Dim TextoInicial As String
        Dim NewRow As DataRow

        Try

            barra = 0

            LM = New Local_M
            dtt = LM.Filter
            FM = New FORMATO_M
            dttIns = FM.AddNew

            For Each row As DataRow In dtt.Rows
                CODIGOPADRE = ""
                CodigoHijo = ""
                Medicion = ""
                Tamaño = Len(row.Item("Descripcion"))
                texto = Mid(row.Item("Descripcion"), 4, Tamaño)
                TextoInicial = texto
                While Medicion = ""
                    barra = InStr(texto, "\")
                    If CODIGOPADRE = "" Then
                        CODIGOPADRE = Replace(Mid(texto, 1, barra - 1), "#", "@")
                        texto = Mid(texto, barra + 1, Tamaño)
                    End If
                    If CodigoHijo = "" Then
                        barra = InStr(texto, "|")
                        CodigoHijo = Replace(Mid(texto, 1, barra - 1), "#", "@")

                        texto = Mid(texto, barra + 1, Tamaño)
                        barra = InStr(texto, "|")
                        texto = Mid(texto, barra + 1, Tamaño)
                    End If
                    If Medicion = "" Then
                        barra = InStr(texto, "|")
                        Medicion = Replace(Mid(texto, 1, barra - 1), ".", ",")
                        texto = Mid(texto, barra + 1, Tamaño)
                    End If

                    If texto = "|" Then
                        NewRow = dttIns.NewRow
                        NewRow.Item("CODIGOPADRE") = CODIGOPADRE
                        NewRow.Item("CodigoHijo") = CodigoHijo
                        NewRow.Item("Medicion") = Medicion
                        NewRow.Item("COMENTARIO") = Replace(Medicion, ",", ".")
                        NewRow.Item("UNIDADES") = 1
                        NewRow.Item("LONGITUD") = 0
                        NewRow.Item("LATITUD") = 0
                        NewRow.Item("Altura") = 0
                        dttIns.Rows.Add(NewRow)
                    Else

                        While texto <> "" And texto <> "|"
                            'Tipo
                            barra = InStr(texto, "\")
                            tipo = Mid(texto, 1, barra - 1)
                            texto = Mid(texto, barra + 1, Tamaño)

                            'Comentario
                            barra = InStr(texto, "\")
                            descripcion = Mid(texto, 1, barra - 1)
                            texto = Mid(texto, barra + 1, Tamaño)
                            'Unidades
                            barra = InStr(texto, "\")
                            If Mid(texto, 1, barra - 1) <> "" Then
                                Cantidad = Replace(Mid(texto, 1, barra - 1), ".", ",")
                            Else
                                Cantidad = 0
                            End If
                            texto = Mid(texto, barra + 1, Tamaño)
                            'Longitud
                            barra = InStr(texto, "\")
                            If Mid(texto, 1, barra - 1) <> "" Then
                                Largo = Replace(Mid(texto, 1, barra - 1), ".", ",")
                            Else
                                Largo = 0
                            End If
                            texto = Mid(texto, barra + 1, Tamaño)
                            'Latitud
                            barra = InStr(texto, "\")
                            If Mid(texto, 1, barra - 1) <> "" Then
                                Anchura = Replace(Mid(texto, 1, barra - 1), ".", ",")
                            Else
                                Anchura = 0
                            End If
                            texto = Mid(texto, barra + 1, Tamaño)
                            'Ancho
                            barra = InStr(texto, "\")
                            If Mid(texto, 1, barra - 1) <> "" Then
                                Altura = Replace(Nz(Mid(texto, 1, barra - 1), 0), ".", ",")
                            Else
                                Altura = 0
                            End If
                            texto = Mid(texto, barra + 1, Tamaño)

                            'Insertamos los valores en la tabla de formato medicion....
                            NewRow = dttIns.NewRow
                            NewRow.Item("CODIGOPADRE") = CODIGOPADRE
                            NewRow.Item("CodigoHijo") = CodigoHijo
                            NewRow.Item("COMENTARIO") = descripcion

                            NewRow.Item("LONGITUD") = Largo
                            NewRow.Item("LATITUD") = Anchura
                            NewRow.Item("Altura") = Altura
                            'If Cantidad = 0 Then Cantidad = 1
                            If Largo = 0 Then Largo = 1
                            If Anchura = 0 Then Anchura = 1
                            If Altura = 0 Then Altura = 1
                            NewRow.Item("Medicion") = Medicion
                            'If Medicion <> 0 And Cantidad = 0 Then
                            '    NewRow.Item("UNIDADES") = 1
                            'Else
                            NewRow.Item("UNIDADES") = Cantidad
                            'End If
                            'If Cantidad = 0 Then
                            '    NewRow.Item("Medicion") = 0
                            'Else
                            '    NewRow.Item("Medicion") = Cantidad * Largo * Anchura * Altura
                            'End If
                            dttIns.Rows.Add(NewRow)

                        End While
                    End If
                End While
            Next

            FM.Update(dttIns)

            dtt = Nothing
            dttIns = Nothing
            LM = Nothing
            FM = Nothing
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Sub FORMATO_T()

        'Damos el formato a los datos de tipo Texto...
        Dim barra As Integer
        Dim Tamaño As Long
        Dim CODIGOPADRE As String
        Dim dtt As DataTable
        Dim dttIns As DataTable
        Dim texto As String
        Dim LT As Local_T
        Dim FT As FORMATO_T
        Dim NewRow As DataRow

        Try

            barra = 0

            LT = New Local_T
            dtt = LT.Filter
            FT = New FORMATO_T
            dttIns = FT.AddNew

            For Each row As DataRow In dtt.Rows
                Tamaño = Len(row.Item("Descripcion"))
                texto = Mid(row.Item("Descripcion"), 4, Tamaño)

                barra = InStr(texto, "|")
                CODIGOPADRE = Mid(texto, 1, barra - 1)
                texto = Replace(Mid(texto, barra + 1, Tamaño - 2), "|", "")


                'Insertamos los valores en la tabla de formato medicion....
                NewRow = dttIns.NewRow
                NewRow.Item("CODIGOCONCEPTO") = Replace(CODIGOPADRE, "#", "@")
                NewRow.Item("texto") = texto
                dttIns.Rows.Add(NewRow)
            Next

            FT.Update(dttIns)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub CrearMateriales(ByRef data As DataTrabajo)

        Dim dtMateriales As DataTable = New BE.DataEngine().Filter("vBC3Materiales", New NumberFilterItem("IDTrabajoPresup", data.IDFiltro), , "IDTrabajoPresup")

        Dim drM As DataRow

        Dim secuencia As Integer
        Dim oFilter As New Filter
        oFilter.Add(New NumberFilterItem("IDTrabajoPresup", FilterOperator.Equal, data.IDFiltro))
        oFilter.Add(New NumberFilterItem("Tipo", FilterOperator.Equal, 3))

        Dim dvlineas As DataView = dtMateriales.DefaultView
        dvlineas.RowFilter = oFilter.Compose(New AdoFilterComposer)
        If dvlineas.Count = 0 Then Exit Sub

        If Not dtMateriales Is Nothing AndAlso dtMateriales.Rows.Count > 0 Then
            For Each drMateriales As DataRow In dtMateriales.Rows
                drM = data.dtMaterialesN.NewRow
                drM("IDLineaMaterial") = AdminData.GetAutoNumeric
                drM("IDTrabajoPresup") = data.IDPadre
                drM("IDPresup") = data.NumPresup
                drM("DescMaterial") = drMateriales("DescTrabajo")
                secuencia = secuencia + 10
                drM("Secuencia") = secuencia

                If drMateriales("Porcentaje") Then
                    drM("IDTrabajoIncremento") = drMateriales("IDTrabajoIncremento")
                    drM("QUnidad") = drMateriales("Factor")
                    drM("QPresup") = drM("QUnidad")
                    'drM("Incremento") = drMateriales("Rendimiento") * 100
                    drM("Incremento") = 0
                    drM("TipoIncremento") = 1
                Else
                    drM("IDMaterial") = drMateriales("IDMaterial")
                    drM("IDUdMedida") = drMateriales("IDUdMedida")
                    If drMateriales("Rendimiento") <> 0 Then
                        drM("QPresup") = drMateriales("Rendimiento")
                    Else
                        drM("QPresup") = 1
                    End If
                    drM("QUnidad") = drM("QPresup")
                    drM("QPresup") = drM("QPresup") * data.cantidad
                    drM("PrecioPresupMatA") = drMateriales("Precio")
                    drM("PrecioVentaA") = drMateriales("Precio")
                    drM("ImpPresupMatA") = xRound(drMateriales("Precio") * drM("QUnidad"), 2)
                    drM("ImpPresupMatVentaA") = xRound(drMateriales("Precio") * drM("QUnidad"), 2)
                End If

                data.dtMaterialesN.Rows.Add(drM.ItemArray)
            Next
        End If
    End Sub
    Public Sub CrearCentros(ByRef data As DataTrabajo)
        Dim dtCentroTrabajo As DataTable = New BE.DataEngine().Filter("vBC3Centros", New NumberFilterItem("IDTrabajoPresup", data.IDFiltro), , "IDTrabajoPresup")

        Dim drOpc As DataRow
        Dim secuencia As Integer

        Dim oFilter As New Filter
        oFilter.Add(New NumberFilterItem("IDTrabajoPresup", FilterOperator.Equal, data.IDFiltro))
        oFilter.Add(New NumberFilterItem("Tipo", FilterOperator.Equal, 2))
        Dim strWhere As String
        strWhere = oFilter.Compose(New AdoFilterComposer)
        Dim dvlineas As DataView = dtCentroTrabajo.DefaultView
        dvlineas.RowFilter = strWhere
        If dvlineas.Count = 0 Then Exit Sub

        If Not dtCentroTrabajo Is Nothing AndAlso dtCentroTrabajo.Rows.Count > 0 Then
            For Each drCentroTrabajo As DataRow In dtCentroTrabajo.Rows

                drOpc = data.dtCentroTrabajoN.NewRow

                drOpc("IDLineaCentro") = AdminData.GetAutoNumeric
                drOpc("IDTrabajoPresup") = data.IDPadre
                drOpc("IDPresup") = data.NumPresup
                secuencia = secuencia + 10
                drOpc("Secuencia") = secuencia
                drOpc("DescCentro") = drCentroTrabajo("DescTrabajo")

                If drCentroTrabajo("Porcentaje") Then
                    drOpc("IDTrabajoIncremento") = drCentroTrabajo("IDTrabajoIncremento")
                    drOpc("HorasUnidad") = drCentroTrabajo("Factor")
                    drOpc("HorasPresupCentro") = drCentroTrabajo("Factor")
                    'drOpc("Incremento") = drCentroTrabajo("Rendimiento") * 100
                    drOpc("Incremento") = 0
                    drOpc("TipoIncremento") = 1
                Else
                    drOpc("IDCentro") = Left(drCentroTrabajo("IDMaterial"), 25)
                    drOpc("HorasUnidad") = drCentroTrabajo("Rendimiento")
                    If drOpc("HorasUnidad") <> 0 Then
                        drOpc("HorasUnidad") = drCentroTrabajo("Rendimiento")
                    Else
                        drOpc("HorasUnidad") = 1
                    End If
                    drOpc("HorasPresupCentro") = drOpc("HorasUnidad") * data.cantidad
                    drOpc("TasaPresupCentroA") = drCentroTrabajo("Precio")
                    drOpc("PrecioVentaA") = drCentroTrabajo("Precio")
                    drOpc("ImpPresupCentroA") = xRound(drCentroTrabajo("Precio") * drOpc("HorasUnidad"), 2)
                    drOpc("ImpPresupCentroVentaA") = xRound(drCentroTrabajo("Precio") * drOpc("HorasUnidad"), 2)
                End If
                data.dtCentroTrabajoN.Rows.Add(drOpc.ItemArray)
            Next
        End If
    End Sub
    Public Sub CrearVarios(ByRef data As DataTrabajo)
        Dim dtVariosTrabajo As DataTable = New BE.DataEngine().Filter("BC3Varios", New NumberFilterItem("IDTrabajoPresup", data.IDFiltro))

        Dim drOpc As DataRow

        Dim oFilter As New Filter
        oFilter.Add(New NumberFilterItem("IDTrabajoPresup", FilterOperator.Equal, data.IDFiltro))
        Dim strWhere As String
        strWhere = oFilter.Compose(New AdoFilterComposer)
        Dim dvlineas As DataView = dtVariosTrabajo.DefaultView
        dvlineas.RowFilter = strWhere
        If dvlineas.Count = 0 Then Exit Sub

        If Not dtVariosTrabajo Is Nothing AndAlso dtVariosTrabajo.Rows.Count > 0 Then
            For Each drVariosTrabajo As DataRow In dtVariosTrabajo.Rows
                drOpc = data.dtVariosTrabajoN.NewRow
                drOpc("IDLineaVarios") = AdminData.GetAutoNumeric
                drOpc("IDTrabajoPresup") = data.IDPadre
                drOpc("IDPresup") = data.NumPresup
                drOpc("IDVarios") = drVariosTrabajo("codigo")
                drOpc("DescVarios") = drVariosTrabajo("resumen")
                drOpc("ImpPresupVariosA") = xRound(drVariosTrabajo("precio") * drVariosTrabajo("rendimiento"), 2)
                drOpc("ImpPresupVariosVentaA") = xRound(drVariosTrabajo("precio") * drVariosTrabajo("rendimiento"), 2)
                data.dtVariosTrabajoN.Rows.Add(drOpc.ItemArray)
            Next
        End If
    End Sub
    Public Sub CrearManoObra(ByRef data As DataTrabajo)
        Dim dtManoObra As DataTable = New BE.DataEngine().Filter("vBC3ManoObra", New NumberFilterItem("IDTrabajoPresup", data.IDFiltro), , "IDTrabajoPresup")
        Dim drOpm As DataRow
        Dim secuencia As Integer
        Dim oFilter As New Filter
        oFilter.Add(New NumberFilterItem("IDTrabajoPresup", FilterOperator.Equal, data.IDFiltro))
        oFilter.Add(New NumberFilterItem("Tipo", FilterOperator.Equal, 1))
        Dim strWhere As String
        strWhere = oFilter.Compose(New AdoFilterComposer)
        Dim dvlineas As DataView = dtManoObra.DefaultView
        dvlineas.RowFilter = strWhere
        If dvlineas.Count = 0 Then Exit Sub

        If Not dtManoObra Is Nothing AndAlso dtManoObra.Rows.Count > 0 Then

            For Each drManoObra As DataRow In dtManoObra.Rows

                drOpm = data.dtManoObraN.NewRow
                drOpm("IDLineaMod") = AdminData.GetAutoNumeric
                drOpm("IDTrabajoPresup") = data.IDPadre
                drOpm("IDPresup") = data.NumPresup
                secuencia = secuencia + 10
                drOpm("Secuencia") = secuencia
                drOpm("DescCategoria") = drManoObra("DescCategoria")

                If drManoObra("Porcentaje") Then
                    drOpm("IDTrabajoIncremento") = drManoObra("IDTrabajoIncremento")
                    drOpm("HorasUnidad") = drManoObra("Factor")
                    drOpm("HorasPresupMod") = drManoObra("Factor")
                    drOpm("IDHora") = data.strTipoHora
                    'drOpm("Incremento") = drManoObra("Rendimiento") * 100
                    drOpm("Incremento") = 0
                    drOpm("TipoIncremento") = 1
                Else
                    drOpm("IDCategoria") = Left(drManoObra("IDCategoria"), 10)
                    drOpm("HorasUnidad") = drManoObra("Rendimiento")
                    If drOpm("HorasUnidad") <> 0 Then
                        drOpm("HorasUnidad") = drManoObra("Rendimiento")
                    Else
                        drOpm("HorasUnidad") = 1
                    End If
                    drOpm("IDHora") = data.strTipoHora
                    drOpm("HorasPresupMod") = drOpm("HorasUnidad") * data.cantidad
                    drOpm("TasaPresupModA") = drManoObra("Precio")
                    drOpm("PrecioVentaA") = drManoObra("Precio")
                    drOpm("ImpPresupModA") = xRound(drManoObra("Precio") * drOpm("HorasUnidad"), 2)
                    drOpm("ImpPresupModVentaA") = xRound(drManoObra("Precio") * drOpm("HorasUnidad"), 2)
                End If

                data.dtManoObraN.Rows.Add(drOpm.ItemArray)
            Next
        End If
    End Sub
    Public Sub CrearMediciones(ByRef data As DataTrabajo)
        Dim f As New Filter
        f.Add(New NumberFilterItem("IdTrabajoPresup", data.IDPadre))
        Dim dtMediciones As DataTable = New BE.DataEngine().Filter("vBC3Mediciones", f, , "IDTrabajoPresup")

        Dim drMed As DataRow
        Dim QPI, Largo, Ancho, Alto As Double
        Dim intReg As Integer

        If Not dtMediciones Is Nothing AndAlso dtMediciones.Rows.Count > 0 Then
            intReg = dtMediciones.Rows.Count
            For Each drMediciones As DataRow In dtMediciones.Rows
                drMed = data.dtMedicionesN.NewRow
                drMed("IdMedicion") = drMediciones("IdMedicion")
                drMed("IDTrabajoPresup") = data.IDPadre
                drMed("IDPresup") = data.NumPresup
                drMed("QPI") = drMediciones("QPI")
                drMed("Largo") = drMediciones("Largo")
                drMed("Ancho") = drMediciones("Ancho")
                drMed("Alto") = drMediciones("Alto")

                QPI = drMediciones("QPI")

                If drMed("Largo") = 0 Then
                    Largo = 1
                Else
                    Largo = drMediciones("Largo")
                End If
                If drMed("Ancho") = 0 Then
                    Ancho = 1
                Else
                    Ancho = drMediciones("Ancho")
                End If
                If drMed("Alto") = 0 Then
                    Alto = 1
                Else
                    Alto = drMediciones("Alto")
                End If
                If intReg = 1 Then
                    drMed("Total") = drMediciones("Total")
                Else
                    drMed("Total") = xRound(QPI * Largo * Alto * Ancho, 2)
                End If

                drMed("DescMedicion") = Nz(drMediciones("DescMedicion"), "  ")

                data.dtMedicionesN.Rows.Add(drMed.ItemArray)
            Next
        End If
    End Sub

    'Public Sub CrearMateriales(ByVal NumPresup As Integer, ByVal CodigoPadre As String, ByVal IDTrabajo As Integer, _
    '                           ByVal dtMaterialesN As DataTable, ByVal QPrevTrabajo As Double)

    '    Dim dtMateriales As DataTable = New BE.DataEngine().Filter("vBC3Materiales", New NumberFilterItem("IDTrabajoPresup", IDTrabajo), , "IDTrabajoPresup")

    '    Dim drM As DataRow

    '    Dim secuencia As Integer
    '    Dim oFilter As New Filter
    '    oFilter.Add(New StringFilterItem("IDTrabajoPresup", FilterOperator.Equal, IDTrabajo))
    '    oFilter.Add(New StringFilterItem("Tipo", FilterOperator.Equal, 3))

    '    Dim dvlineas As DataView = dtMateriales.DefaultView
    '    dvlineas.RowFilter = oFilter.Compose(New AdoFilterComposer)
    '    If dvlineas.Count = 0 Then Exit Sub

    '    If Not dtMateriales Is Nothing AndAlso dtMateriales.Rows.Count > 0 Then
    '        For Each drMateriales As DataRow In dtMateriales.Rows
    '            drM = dtMaterialesN.NewRow
    '            drM("IDLineaMaterial") = AdminData.GetAutoNumeric
    '            drM("IDTrabajoPresup") = IDTrabajo
    '            drM("IDPresup") = NumPresup
    '            drM("DescMaterial") = drMateriales("DescTrabajo")
    '            secuencia = secuencia + 10
    '            drM("Secuencia") = secuencia

    '            If drMateriales("Porcentaje") Then
    '                drM("IDTrabajoIncremento") = drMateriales("IDTrabajoIncremento")
    '                drM("QUnidad") = drMateriales("Factor")
    '                drM("QPresup") = drM("QUnidad")
    '                drM("Incremento") = drMateriales("Rendimiento") * 100
    '                drM("TipoIncremento") = 1
    '            Else
    '                drM("IDMaterial") = drMateriales("IDMaterial")
    '                drM("IDUdMedida") = drMateriales("IDUdMedida")
    '                If drMateriales("Rendimiento") <> 0 Then
    '                    drM("QPresup") = drMateriales("Rendimiento")
    '                Else
    '                    drM("QPresup") = 1
    '                End If
    '                drM("QUnidad") = drM("QPresup")
    '                drM("QPresup") = drM("QPresup") * QPrevTrabajo
    '                drM("PrecioPresupMatA") = drMateriales("Precio")
    '                drM("PrecioVentaA") = drMateriales("Precio")
    '                drM("ImpPresupMatA") = xRound(drMateriales("Precio") * drM("QUnidad"), 2)
    '                drM("ImpPresupMatVentaA") = xRound(drMateriales("Precio") * drM("QUnidad"), 2)
    '            End If

    '            dtMaterialesN.Rows.Add(drM.ItemArray)
    '        Next
    '    End If
    'End Sub

    'Public Sub CrearCentros(ByVal NumPresup As Integer, ByVal CodigoPadre As String, ByVal IDTrabajo As Integer, ByVal dtCentroTrabajoN As DataTable, ByVal QPrevTrabajo As Double)
    '    Dim dtCentroTrabajo As DataTable = New BE.DataEngine().Filter("BC3Centros", New StringFilterItem("CodigoPadre", CodigoPadre), , "IDTrabajoPresup")

    '    Dim drOpc As DataRow
    '    Dim secuencia As Integer

    '    Dim oFilter As New Filter
    '    oFilter.Add(New StringFilterItem("CodigoPadre", FilterOperator.Equal, CodigoPadre))
    '    oFilter.Add(New StringFilterItem("Tipo", FilterOperator.Equal, 2))
    '    Dim strWhere As String
    '    strWhere = oFilter.Compose(New AdoFilterComposer)
    '    Dim dvlineas As DataView = dtCentroTrabajo.DefaultView
    '    dvlineas.RowFilter = strWhere
    '    If dvlineas.Count = 0 Then Exit Sub

    '    If Not dtCentroTrabajo Is Nothing AndAlso dtCentroTrabajo.Rows.Count > 0 Then
    '        For Each drCentroTrabajo As DataRow In dtCentroTrabajo.Rows

    '            drOpc = dtCentroTrabajoN.NewRow

    '            drOpc("IDLineaCentro") = AdminData.GetAutoNumeric
    '            drOpc("IDTrabajoPresup") = IDTrabajo
    '            drOpc("IDPresup") = NumPresup
    '            secuencia = secuencia + 10
    '            drOpc("Secuencia") = secuencia
    '            drOpc("DescCentro") = drCentroTrabajo("DescTrabajo")

    '            If drCentroTrabajo("Porcentaje") Then
    '                drOpc("IDTrabajoIncremento") = drCentroTrabajo("IDTrabajoIncremento")
    '                drOpc("HorasUnidad") = drCentroTrabajo("Factor")
    '                drOpc("HorasPresupCentro") = drCentroTrabajo("Factor")
    '                drOpc("Incremento") = drCentroTrabajo("Rendimiento") * 100
    '                drOpc("TipoIncremento") = 1
    '            Else
    '                drOpc("IDCentro") = Left(drCentroTrabajo("IDMaterial"), 25)
    '                drOpc("HorasUnidad") = drCentroTrabajo("Rendimiento")
    '                If drOpc("HorasUnidad") <> 0 Then
    '                    drOpc("HorasUnidad") = drCentroTrabajo("Rendimiento")
    '                Else
    '                    drOpc("HorasUnidad") = 1
    '                End If
    '                drOpc("HorasPresupCentro") = drOpc("HorasUnidad") * QPrevTrabajo
    '                drOpc("TasaPresupCentroA") = drCentroTrabajo("Precio")
    '                drOpc("PrecioVentaA") = drCentroTrabajo("Precio")
    '                drOpc("ImpPresupCentroA") = xRound(drCentroTrabajo("Precio") * drOpc("HorasUnidad"), 2)
    '                drOpc("ImpPresupCentroVentaA") = xRound(drCentroTrabajo("Precio") * drOpc("HorasUnidad"), 2)
    '            End If
    '            dtCentroTrabajoN.Rows.Add(drOpc.ItemArray)
    '        Next
    '    End If
    'End Sub
    'Public Sub CrearVarios(ByVal NumPresup As Integer, ByVal CodigoPadre As String, ByVal IDTrabajo As Integer, ByVal dtVariosTrabajoN As DataTable)
    '    Dim dtVariosTrabajo As DataTable = New BE.DataEngine().Filter("BC3Varios", New StringFilterItem("CodigoPadre", CodigoPadre))

    '    Dim drOpc As DataRow

    '    Dim oFilter As New Filter
    '    oFilter.Add(New StringFilterItem("CodigoPadre", FilterOperator.Equal, CodigoPadre))
    '    Dim strWhere As String
    '    strWhere = oFilter.Compose(New AdoFilterComposer)
    '    Dim dvlineas As DataView = dtVariosTrabajo.DefaultView
    '    dvlineas.RowFilter = strWhere
    '    If dvlineas.Count = 0 Then Exit Sub

    '    If Not dtVariosTrabajo Is Nothing AndAlso dtVariosTrabajo.Rows.Count > 0 Then
    '        For Each drVariosTrabajo As DataRow In dtVariosTrabajo.Rows
    '            drOpc = dtVariosTrabajoN.NewRow
    '            drOpc("IDLineaVarios") = AdminData.GetAutoNumeric
    '            drOpc("IDTrabajoPresup") = IDTrabajo
    '            drOpc("IDPresup") = NumPresup
    '            drOpc("IDVarios") = drVariosTrabajo("codigo")
    '            drOpc("DescVarios") = drVariosTrabajo("resumen")
    '            drOpc("ImpPresupVariosA") = xRound(drVariosTrabajo("precio") * drVariosTrabajo("rendimiento"), 2)
    '            drOpc("ImpPresupVariosVentaA") = xRound(drVariosTrabajo("precio") * drVariosTrabajo("rendimiento"), 2)
    '            dtVariosTrabajoN.Rows.Add(drOpc.ItemArray)
    '        Next
    '    End If
    'End Sub
    Sub Formato_D_ID()
        Dim FD As New Formato_D

        AdminData.Execute("UPDATE FORMATO_D SET IDPadre =FORMATO_C.id FROM FORMATO_D LEFT OUTER JOIN FORMATO_C ON FORMATO_D.CODIGOPADRE = FORMATO_C.CODIGO")
        AdminData.Execute("UPDATE FORMATO_D SET IDHijo =FORMATO_C.id FROM FORMATO_D LEFT OUTER JOIN FORMATO_C ON FORMATO_D.CODIGOHIJO = FORMATO_C.CODIGO")

        Dim dtFormatoD As DataTable = AdminData.Filter("vBC3CodigoRepetido")
        Dim dtRepetidos As DataTable
        Dim IDHijo As Integer
        Dim dttIns As DataTable
        Dim NewRow As DataRow

        If Not dtFormatoD Is Nothing AndAlso dtFormatoD.Rows.Count > 0 Then
            dttIns = New Formato_C().AddNew
            For Each drFormatoD As DataRow In dtFormatoD.Rows

                Dim f As New Filter
                f.Add(New StringFilterItem("IDHijo", FilterOperator.Equal, drFormatoD("IDHijo")))
                dtRepetidos = New BE.DataEngine().Filter("FORMATO_D", f, , "ID")
                If Not dtRepetidos Is Nothing AndAlso dtRepetidos.Rows.Count > 0 Then
                    IDHijo = dtFormatoD.Rows(0)("IDHijo")
                    Dim drFormatoC As DataRow = New Formato_C().GetItemRow(dtRepetidos.Rows(0)("IDHijo"))
                    For Each drRepetidos As DataRow In dtRepetidos.Rows
                        If IDHijo <> drRepetidos("IDHijo") Then
                            IDHijo = 0
                            NewRow = dttIns.NewRow
                            Dim ID As Integer = AdminData.GetAutoNumeric
                            NewRow("ID") = ID
                            NewRow("codigo") = drFormatoC("codigo")
                            NewRow("porcentaje") = drFormatoC("porcentaje")
                            NewRow("unidad") = drFormatoC("unidad")
                            NewRow("precio") = drFormatoC("precio")
                            NewRow("Resumen") = drFormatoC("Resumen")
                            NewRow("fecha") = drFormatoC("fecha")
                            NewRow("tipo") = drFormatoC("tipo")
                            NewRow("Capitulo") = drFormatoC("Capitulo")
                            NewRow("Nivel") = drFormatoC("Nivel")
                            NewRow("CODIGOINICIAL") = String.Empty
                            dttIns.Rows.Add(NewRow)
                            drRepetidos("IDHijo") = ID
                        End If
                IDHijo = 0
                    Next
                   
                    FD.Update(dtRepetidos)
                End If

            Next
            Dim FC As New Formato_C
            FC.Update(dttIns)
        End If


    End Sub
    Sub Formato_M_ID()
        Dim FD As New Formato_D

        AdminData.Execute("UPDATE FORMATO_M SET IDPadre =FORMATO_C.id FROM FORMATO_M LEFT OUTER JOIN FORMATO_C ON FORMATO_M.CODIGOPADRE = FORMATO_C.CODIGO")
        AdminData.Execute("UPDATE FORMATO_M SET IDHijo =FORMATO_C.id FROM FORMATO_M LEFT OUTER JOIN FORMATO_C ON FORMATO_M.CODIGOHIJO = FORMATO_C.CODIGO")


    End Sub
    'Public Sub CrearMediciones(ByVal NumPresup As Integer, ByVal CodigoPadre As String, ByVal IDTrabajo As Integer, ByVal dtMedicionesN As DataTable)
    '    Dim f As New Filter
    '    f.Add(New StringFilterItem("CodigoHijo", CodigoPadre))
    '    f.Add(New NumberFilterItem("IdTrabajoPresup", IDTrabajo))
    '    Dim dtMediciones As DataTable = New BE.DataEngine().Filter("BC3Mediciones", f, , "IDTrabajoPresup")

    '    Dim drMed As DataRow
    '    Dim QPI, Largo, Ancho, Alto As Double
    '    Dim intReg As Integer

    '    If Not dtMediciones Is Nothing AndAlso dtMediciones.Rows.Count > 0 Then
    '        intReg = dtMediciones.Rows.Count
    '        For Each drMediciones As DataRow In dtMediciones.Rows
    '            drMed = dtMedicionesN.NewRow
    '            drMed("IdMedicion") = drMediciones("IdMedicion")
    '            drMed("IDTrabajoPresup") = IDTrabajo
    '            drMed("IDPresup") = NumPresup
    '            drMed("QPI") = drMediciones("QPI")
    '            drMed("Largo") = drMediciones("Largo")
    '            drMed("Ancho") = drMediciones("Ancho")
    '            drMed("Alto") = drMediciones("Alto")

    '            QPI = drMediciones("QPI")

    '            If drMed("Largo") = 0 Then
    '                Largo = 1
    '            Else
    '                Largo = drMediciones("Largo")
    '            End If
    '            If drMed("Ancho") = 0 Then
    '                Ancho = 1
    '            Else
    '                Ancho = drMediciones("Ancho")
    '            End If
    '            If drMed("Alto") = 0 Then
    '                Alto = 1
    '            Else
    '                Alto = drMediciones("Alto")
    '            End If
    '            If intReg = 1 Then
    '                drMed("Total") = drMediciones("Total")
    '            Else
    '                drMed("Total") = xRound(QPI * Largo * Alto * Ancho, 2)
    '            End If

    '            drMed("DescMedicion") = Nz(drMediciones("DescMedicion"), "  ")

    '            dtMedicionesN.Rows.Add(drMed.ItemArray)
    '        Next
    '    End If
    'End Sub
    'Public Sub CrearManoObra(ByVal NumPresup As Integer, ByVal CodigoPadre As String, ByVal strTipoHora As String, ByVal IDTrabajo As Integer, ByVal dtManoObraN As DataTable, ByVal QPrevTrabajo As Double)
    '    Dim dtManoObra As DataTable = New BE.DataEngine().Filter("BC3ManoObra", New StringFilterItem("CodigoPadre", CodigoPadre), , "IDTrabajoPresup")
    '    Dim drOpm As DataRow
    '    Dim secuencia As Integer
    '    Dim oFilter As New Filter
    '    oFilter.Add(New StringFilterItem("CodigoPadre", FilterOperator.Equal, CodigoPadre))
    '    oFilter.Add(New StringFilterItem("Tipo", FilterOperator.Equal, 1))
    '    Dim strWhere As String
    '    strWhere = oFilter.Compose(New AdoFilterComposer)
    '    Dim dvlineas As DataView = dtManoObra.DefaultView
    '    dvlineas.RowFilter = strWhere
    '    If dvlineas.Count = 0 Then Exit Sub

    '    If Not dtManoObra Is Nothing AndAlso dtManoObra.Rows.Count > 0 Then

    '        For Each drManoObra As DataRow In dtManoObra.Rows

    '            drOpm = dtManoObraN.NewRow
    '            drOpm("IDLineaMod") = AdminData.GetAutoNumeric
    '            drOpm("IDTrabajoPresup") = IDTrabajo
    '            drOpm("IDPresup") = NumPresup
    '            secuencia = secuencia + 10
    '            drOpm("Secuencia") = secuencia
    '            drOpm("DescCategoria") = drManoObra("DescCategoria")

    '            If drManoObra("Porcentaje") Then
    '                drOpm("IDTrabajoIncremento") = drManoObra("IDTrabajoIncremento")
    '                drOpm("HorasUnidad") = drManoObra("Factor")
    '                drOpm("HorasPresupMod") = drManoObra("Factor")
    '                drOpm("IDHora") = strTipoHora
    '                drOpm("Incremento") = drManoObra("Rendimiento") * 100
    '                drOpm("TipoIncremento") = 1
    '            Else
    '                drOpm("IDCategoria") = Left(drManoObra("IDCategoria"), 10)
    '                drOpm("HorasUnidad") = drManoObra("Rendimiento")
    '                If drOpm("HorasUnidad") <> 0 Then
    '                    drOpm("HorasUnidad") = drManoObra("Rendimiento")
    '                Else
    '                    drOpm("HorasUnidad") = 1
    '                End If
    '                drOpm("IDHora") = strTipoHora
    '                drOpm("HorasPresupMod") = drOpm("HorasUnidad") * QPrevTrabajo
    '                drOpm("TasaPresupModA") = drManoObra("Precio")
    '                drOpm("PrecioVentaA") = drManoObra("Precio")
    '                drOpm("ImpPresupModA") = xRound(drManoObra("Precio") * drOpm("HorasUnidad"), 2)
    '                drOpm("ImpPresupModVentaA") = xRound(drManoObra("Precio") * drOpm("HorasUnidad"), 2)
    '            End If

    '            dtManoObraN.Rows.Add(drOpm.ItemArray)
    '        Next
    '    End If
    'End Sub
    Public Function CrearTipoPresupuesto(ByVal NumPresup As Integer) As String
        Dim CodigoTipo As String = String.Empty
        Dim fc As New Formato_C
        Dim dtfc As DataTable = fc.Filter(New LikeFilterItem("codigo", "%@@"), , "Codigo, Resumen")
        Dim dtot As DataTable
        Dim ot As New Obra.ObraTipo
        Dim dtotn As DataTable = ot.AddNew
        Dim drOt As DataRow

        If Not dtfc Is Nothing AndAlso dtfc.Rows.Count > 0 Then
            dtotn = ot.AddNew
            For Each drfc As DataRow In dtfc.Rows
                CodigoTipo = Left(Replace(drfc("Codigo"), "@@", ""), 10)
                dtot = ot.Filter(New StringFilterItem("idTipoObra", CodigoTipo), , "IDTipoObra")
                If dtot.Rows.Count = 0 Then

                    drOt = dtotn.NewRow
                    If Len(CodigoTipo) <= 10 Then
                        drOt("IDTipoObra") = CodigoTipo
                    End If
                    If Length(drfc("Resumen")) > 0 Then
                        drOt("DescTipoObra") = Replace(drfc("Resumen"), "@@", "")
                    End If
                    dtotn.Rows.Add(drOt.ItemArray)

                End If
            Next
        End If

        AdminData.Execute("UPDATE formato_c SET codigo =  replace(codigo,'@','') WHERE codigo in (SELECT CODIGOHIJO  + '@' FROM formato_d)")
        AdminData.Execute("UPDATE formato_d SET CODIGOHIJO =  replace(CODIGOHIJO,'@','') WHERE  CODIGOHIJO LIKE '%@'")
        AdminData.Execute("UPDATE formato_d SET CODIGOPADRE =  replace(CODIGOPADRE,'@','') WHERE  CODIGOPADRE LIKE '%@'")
        AdminData.Execute("UPDATE formato_M SET CODIGOHIJO =  replace(CODIGOHIJO,'@','') WHERE  CODIGOHIJO LIKE '%@'")
        AdminData.Execute("UPDATE formato_M SET CODIGOPADRE =  replace(CODIGOPADRE,'@','') WHERE  CODIGOPADRE LIKE '%@'")
        Formato_D_ID()
        Formato_M_ID()
        BusinessHelper.UpdatePackage(New UpdatePackage(dtotn.TableName, dtotn))

        AdminData.Execute("UPDATE FORMATO_C SET IdTipoObra ='" & CodigoTipo & "'")

        AdminData.Execute("UPDATE tbObraPresupCabecera set IDTipoObra='" & CodigoTipo & "' where IDPresup=" & NumPresup)

        Return CodigoTipo
    End Function
    Public Sub CrearUnidadesMedida(Optional ByVal blnTrabajo As Boolean = True)
        Dim dtUd As DataTable
        Dim um As New Negocio.UdMedida
        Dim dtum As DataTable
        Dim dtumn As DataTable

        If blnTrabajo Then
            dtUd = New BE.DataEngine().Filter("formato_c", New StringFilterItem("Tipo", FilterOperator.NotEqual, 3), "DISTINCT UNIDAD")
        Else
            dtUd = New BE.DataEngine().Filter("formato_c", "DISTINCT UNIDAD", "")
        End If

        If Not dtUd Is Nothing AndAlso dtUd.Rows.Count > 0 Then
            dtumn = um.AddNew
            For Each drUd As DataRow In dtUd.Rows
                dtum = um.SelOnPrimaryKey(drUd("unidad"))
                If dtum.Rows.Count = 0 Then
                    Dim drUm As DataRow = dtumn.NewRow
                    drUm("IDUDMedida") = drUd("unidad")
                    drUm("DescUdMedida") = drUd("unidad")
                    dtumn.Rows.Add(drUm.ItemArray)
                End If
            Next
        End If
        BusinessHelper.UpdatePackage(New UpdatePackage(dtumn.TableName, dtumn))
    End Sub

    Public Sub CrearTrabajos(ByVal NumPresup As Integer, ByVal CodigoPadre As String, _
                             Optional ByVal blnImportarMateriales As Boolean = True, _
                             Optional ByVal blnCrearMateriales As Boolean = True, _
                             Optional ByVal blnImportarMOD As Boolean = True, _
                             Optional ByVal blnImportarCentros As Boolean = True, _
                             Optional ByVal blnImportarMediciones As Boolean = True)

        Dim Orden As Integer

        Orden = 10
        Dim dtTrabajos As DataTable
        Dim data As New DataTrabajo
        Dim t As New Obra.ObraTrabajoPresup
        Dim dtTrabajosN As DataTable = t.AddNew
        Dim pa As New Parametro
        Dim dblAcumular As Boolean
        Dim strIDArticulo As String
        Dim IDTrabajo As Integer
        Dim ObraMaterial As New ObraPresupMaterial
        Dim ObraMod As New ObraPresupMOD
        Dim ObraVarios As New ObraPresupVarios
        Dim ObraCentro As New ObraPresupCentro
        Dim ObraMedicion As New ObraPresupMedicion

        Dim strTipoHora As String = pa.HoraPred

        Me.BeginTx()


        Dim Nivel As Integer

        strIDArticulo = pa.ArticuloFacturacionProyectos()
        dblAcumular = pa.NoAcumularEnTrabajo()
        Dim drT As DataRow
        'Trabajos de tipo Porcentaje
        Dim fc As New Formato_C
        Dim fPor As New Filter


        fPor.Add(New BooleanFilterItem("Porcentaje", FilterOperator.Equal, True))
        dtTrabajos = fc.Filter(fPor)

        data.NumPresup = NumPresup
        data.IDTipoObra = CodigoPadre
        data.blnImportarCentros = blnImportarCentros
        data.blnImportarMateriales = blnImportarMateriales
        data.blnImportarMediciones = blnImportarMediciones
        data.blnImportarMOD = blnImportarMOD
        Dim ObraTrabajo As New ObraTrabajoPresup

        data.dtTrabajosN = ObraTrabajo.AddNew

        Dim ObraTipoTrabajo As New ObraTipoTrabajo
        data.dtObraTipoTrabajoN = ObraTipoTrabajo.AddNew

        Dim drObraTipoTrabajo As DataRow
        Dim dtExiste As DataTable

        Dim ObraSubTipoTrabajo As New ObraSubtipoTrabajo
        data.dtObraSubTipoTrabajoN = ObraSubTipoTrabajo.AddNew


        Dim ObraSubSubTipoTrabajo As New ObraSubSubtipoTrabajo
        data.dtObraSubSubTipoTrabajoN = ObraSubSubTipoTrabajo.AddNew


        data.dtMaterialesN = ObraMaterial.AddNew
        data.dtManoObraN = ObraMod.AddNew
        data.dtCentroTrabajoN = ObraCentro.AddNew
        data.dtMedicionesN = ObraMedicion.AddNew
        data.dtVariosTrabajoN = ObraVarios.AddNew
        data.dblAcumular = dblAcumular

        If Not dtTrabajos Is Nothing AndAlso dtTrabajos.Rows.Count > 0 Then
            Dim IDUDMedida As String = New Parametro().UdMedidaPred
            For Each drTrabajo As DataRow In dtTrabajos.Rows
                drT = data.dtTrabajosN.NewRow
                drT("IDTrabajoPresup") = drTrabajo("ID")
                IDTrabajo = drT("IDTrabajoPresup")
                drT("IDPresup") = NumPresup
                drT("IDTipoObra") = CodigoPadre
                drT("CodTrabajo") = drTrabajo("CODIGO")
                drT("Secuencia") = Orden
                drT("DescTrabajo") = drTrabajo("RESUMEN")
                drT("IDArticulo") = strIDArticulo
                drT("Incremento") = drTrabajo("Precio")
                drT("IDUDMedida") = IDUDMedida
                drT("NoAcumular") = data.dblAcumular
                drT("Nivel") = 0
                drT("TipoLinea") = enumTipoLineaTrabajo.tltPorcentajeConcepto
                Orden = Orden + 10
                data.dtTrabajosN.Rows.Add(drT.ItemArray)
                drTrabajo("IDTrabajoIncremento") = drT("IDTrabajoPresup")
            Next
        End If

        dtTrabajos = New BE.DataEngine().Filter("vBC3CapitulosPadres", "", "CODIGOINICIAL not like '%@@%'", "IDTrabajoPresup")

        If Not dtTrabajos Is Nothing AndAlso dtTrabajos.Rows.Count > 0 Then
            For Each drTrabajo As DataRow In dtTrabajos.Rows
                Dim f As New Filter
                f.Add(New StringFilterItem("IdTipoObra", CodigoPadre))
                f.Add(New StringFilterItem("IDTipoTrabajo", drTrabajo("CodTrabajo")))
                dtExiste = ObraTipoTrabajo.Filter(f)
                If dtExiste.Rows.Count = 0 Then
                    Dim dtObraTipo As DataTable = ObraTipoTrabajo.AddNew
                    drObraTipoTrabajo = data.dtObraTipoTrabajoN.NewRow
                    drObraTipoTrabajo("IDTipoObra") = CodigoPadre
                    drObraTipoTrabajo("IDTipoTrabajo") = drTrabajo("CodTrabajo")
                    drObraTipoTrabajo("DesctipoTrabajo") = drTrabajo("DescTrabajo")
                    dtObraTipo.Rows.Add(drObraTipoTrabajo.ItemArray)
                    BusinessHelper.UpdatePackage(New UpdatePackage(dtObraTipo.TableName, dtObraTipo))
                End If

                drT = dtTrabajosN.NewRow
                drT("IDTrabajoPresup") = drTrabajo("IDTrabajoPresup")
                IDTrabajo = drT("IDTrabajoPresup")
                drT("IDPresup") = NumPresup
                drT("IDTipoObra") = CodigoPadre
                drT("IDTipoTrabajo") = drTrabajo("CodTrabajo")
                drT("CodTrabajo") = drTrabajo("CodTrabajo")
                drT("Secuencia") = Orden
                drT("DescTrabajo") = drTrabajo("DescTrabajo")
                drT("IDArticulo") = strIDArticulo
                drT("IDUdMedida") = drTrabajo("IDUdMedida")
                drT("NoAcumular") = data.dblAcumular
                drT("Nivel") = 0
                drT("TipoLinea") = enumTipoLineaTrabajo.tltCapitulo

                '  Revisar si es capítulo o trabajo 
                Dim dtTrabajosNoPadres As DataTable = New BE.DataEngine().Filter("vBC3CapitulosPadresSinHijos", New NumberFilterItem("IDTrabajoPresup", drTrabajo("IDTrabajoPresup")))
                If Not dtTrabajosNoPadres Is Nothing AndAlso dtTrabajosNoPadres.Rows.Count > 0 Then
                    drT("QPresup") = 1
                    drT("TipoLinea") = enumTipoLineaTrabajo.tltTrabajo
                    drT("ImpPresupTrabajoA") = dtTrabajosNoPadres.Rows(0)("Precio")
                    drT("ImpPresupTrabajoVentaA") = dtTrabajosNoPadres.Rows(0)("Precio")
                    drT("ImpPresupQTrabajoA") = dtTrabajosNoPadres.Rows(0)("Precio")
                    drT("ImpPresupQTrabajoVentaA") = dtTrabajosNoPadres.Rows(0)("Precio")

                    Orden = Orden + 10
                End If

                drT("QPresup") = 1
                '  drT("TipoLinea") = enumTipoLineaTrabajo.tltTrabajo
                'drT("ImpPresupTrabajoA") = drTrabajo("Precio")
                'drT("ImpPresupTrabajoVentaA") = drTrabajo("Precio")
                'drT("ImpPresupQTrabajoA") = drTrabajo("Precio")
                'drT("ImpPresupQTrabajoVentaA") = drTrabajo("Precio")

                Orden = Orden + 10
                data.orden = Orden
                data.strIDArticulo = strIDArticulo
                data.dtTrabajosN.Rows.Add(drT.ItemArray)
                data.cantidad = 1
                data.IDPadre = drTrabajo("IDTrabajoPresup")
                data.IDFiltro = data.IDPadre
                If data.blnImportarMateriales Then CrearMateriales(data)
                If data.blnImportarMOD Then CrearManoObra(data)
                If data.blnImportarCentros Then CrearCentros(data)
                If data.blnImportarMediciones Then CrearMediciones(data)
            Next
        End If
        data.nivel = 1
        CrearNivelesFormatoC(data)
        BusinessHelper.UpdatePackage(New UpdatePackage(data.dtTrabajosN.TableName, data.dtTrabajosN))
        If blnImportarMateriales Then BusinessHelper.UpdatePackage(New UpdatePackage(data.dtMaterialesN.TableName, data.dtMaterialesN))
        If blnImportarMOD Then BusinessHelper.UpdatePackage(New UpdatePackage(data.dtManoObraN.TableName, data.dtManoObraN))
        If blnImportarMediciones Then BusinessHelper.UpdatePackage(New UpdatePackage(data.dtMedicionesN.TableName, data.dtMedicionesN))
        If blnImportarCentros Then BusinessHelper.UpdatePackage(New UpdatePackage(data.dtCentroTrabajoN.TableName, data.dtCentroTrabajoN))
        BusinessHelper.UpdatePackage(New UpdatePackage(data.dtVariosTrabajoN.TableName, data.dtVariosTrabajoN))
    End Sub
    'Public Sub CrearNivelesRecursivos_old(ByVal NumPresup As Integer, ByVal dblAcumular As Boolean, ByVal Nivel As Integer, _
    '                                    ByRef Orden As Integer, ByVal IDTrabajo As Integer, ByVal strIDArticulo As String, ByVal strTipoHora As String, _
    '                                    ByRef dtTrabajosN As DataTable, ByRef dtObraSubTipoTrabajoN As DataTable, ByRef dtObraSubSubTipoTrabajoN As DataTable, _
    '                                    ByRef dtMaterialesN As DataTable, ByRef dtManoObraN As DataTable, ByRef dtCentroTrabajoN As DataTable, ByRef dtMedicionesN As DataTable, _
    '                                    ByVal blnImportarMateriales As Boolean, ByVal blnImportarMOD As Boolean, ByVal blnImportarCentros As Boolean, ByVal blnImportarMediciones As Boolean, _
    '                                    ByVal CodigoPadre As String, ByVal CodTrabajo As String, Optional ByVal CodTrabajoHijo As String = "", Optional ByVal CodTrabajoNieto As String = "")
    '    'Trabajos Dependientes

    '    Dim ObraSubTipoTrabajo As New ObraSubtipoTrabajo
    '    Dim drObraSubTipoTrabajo As DataRow
    '    Dim ObraSubSubTipoTrabajo As New ObraSubSubtipoTrabajo
    '    Dim drObraSubSubTipoTrabajo As DataRow
    '    Dim drT As DataRow
    '    Dim dtExiste As DataTable
    '    Dim dtTrabajos As DataTable
    '    Dim IDTrabajoHijo As Integer
    '    Dim CodTrabajoHijorecursivo As String = String.Empty
    '    Dim CodTrabajoNietorecursivo As String = String.Empty

    '    Nivel = Nivel + 1
    '    If CodTrabajoNieto <> "" Then
    '        dtTrabajos = New BE.DataEngine().Filter("BC3CapitulosHijos", New StringFilterItem("CODIGOPADRE", CodTrabajoNieto), , "IDtrabajoPresup")
    '    ElseIf CodTrabajoHijo <> "" Then
    '        dtTrabajos = New BE.DataEngine().Filter("BC3CapitulosHijos", New StringFilterItem("CODIGOPADRE", CodTrabajoHijo), , "IDtrabajoPresup")
    '    Else
    '        dtTrabajos = New BE.DataEngine().Filter("BC3CapitulosHijos", New StringFilterItem("CODIGOPADRE", CodTrabajo), , "IDtrabajoPresup")
    '    End If

    '    dtTrabajos = New BE.DataEngine().Filter("BC3CapitulosHijos", New NumberFilterItem("IdPadre", IDTrabajo), "", "IDtrabajoPresup")
    '    If Not dtTrabajos Is Nothing AndAlso dtTrabajos.Rows.Count > 0 Then
    '        For Each drTrabajoHijo As DataRow In dtTrabajos.Select
    '            If drTrabajoHijo("CodTrabajo") <> CodigoPadre Then
    '                If CodTrabajoHijo = "" Then
    '                    Dim f As New Filter
    '                    f.Add(New StringFilterItem("IDTipoObra", CodigoPadre))
    '                    f.Add(New StringFilterItem("IDTipoTrabajo", CodTrabajo))
    '                    f.Add(New StringFilterItem("IDSubTipoTrabajo", drTrabajoHijo("CodTrabajo")))
    '                    dtExiste = ObraSubTipoTrabajo.Filter(f)
    '                    If dtExiste.Rows.Count = 0 Then
    '                        Dim dtObraSubTipo As DataTable = ObraSubTipoTrabajo.AddNew

    '                        drObraSubTipoTrabajo = dtObraSubTipoTrabajoN.NewRow
    '                        drObraSubTipoTrabajo("IDTipoObra") = CodigoPadre
    '                        drObraSubTipoTrabajo("IDTipoTrabajo") = CodTrabajo
    '                        drObraSubTipoTrabajo("IDSubTipoTrabajo") = drTrabajoHijo("CodTrabajo")
    '                        drObraSubTipoTrabajo("DescSubtipoTrabajo") = drTrabajoHijo("DescTrabajo")
    '                        dtObraSubTipo.Rows.Add(drObraSubTipoTrabajo.ItemArray)

    '                        BusinessHelper.UpdatePackage(New UpdatePackage(dtObraSubTipo.TableName, dtObraSubTipo))
    '                    End If
    '                End If
    '                If CodTrabajoHijo <> "" And CodTrabajoNieto = "" Then
    '                    Dim f As New Filter
    '                    f.Add(New StringFilterItem("IDTipoObra", CodigoPadre))
    '                    f.Add(New StringFilterItem("IDTipoTrabajo", CodTrabajo))
    '                    f.Add(New StringFilterItem("IDSubTipoTrabajo", CodTrabajoHijo))
    '                    f.Add(New StringFilterItem("IDSubSubTipoTrabajo", drTrabajoHijo("CodTrabajo")))
    '                    dtExiste = ObraSubSubTipoTrabajo.Filter(f)
    '                    If dtExiste.Rows.Count = 0 Then
    '                        Dim dtObraSubSubTipo As DataTable = ObraSubSubTipoTrabajo.AddNew
    '                        drObraSubSubTipoTrabajo = dtObraSubSubTipoTrabajoN.NewRow
    '                        drObraSubSubTipoTrabajo("IDTipoObra") = CodigoPadre
    '                        drObraSubSubTipoTrabajo("IDTipoTrabajo") = CodTrabajo
    '                        drObraSubSubTipoTrabajo("IDSubTipoTrabajo") = CodTrabajoHijo
    '                        drObraSubSubTipoTrabajo("IDSubSubTipoTrabajo") = drTrabajoHijo("CodTrabajo")
    '                        drObraSubSubTipoTrabajo("DescSubSubtipoTrabajo") = drTrabajoHijo("DescTrabajo")
    '                        dtObraSubSubTipo.Rows.Add(drObraSubSubTipoTrabajo.ItemArray)

    '                        BusinessHelper.UpdatePackage(New UpdatePackage(dtObraSubSubTipo.TableName, dtObraSubSubTipo))
    '                    End If
    '                End If
    '                drT = dtTrabajosN.NewRow

    '                drT("IDTrabajoPresup") = drTrabajoHijo("IDTrabajoPresup")
    '                drT("IDPresup") = NumPresup
    '                drT("IDTipoObra") = CodigoPadre
    '                drT("IDTipoTrabajo") = CodTrabajo
    '                drT("IDSubTipoTrabajo") = drTrabajoHijo("CodTrabajo")
    '                drT("CodTrabajo") = drTrabajoHijo("CodTrabajo")
    '                If CodTrabajoHijo <> "" And CodTrabajoNieto = "" Then
    '                    drT("IDSubTipoTrabajo") = CodTrabajoHijo
    '                    drT("SubSubTipoTrabajo") = drTrabajoHijo("CodTrabajo")
    '                End If
    '                If CodTrabajoHijo <> "" And CodTrabajoNieto <> "" Then
    '                    drT("IDSubTipoTrabajo") = CodTrabajoHijo
    '                    drT("SubSubTipoTrabajo") = CodTrabajoNieto
    '                End If
    '                drT("DescTrabajo") = drTrabajoHijo("DescTrabajo")
    '                drT("IDArticulo") = strIDArticulo
    '                drT("IDUdMedida") = drTrabajoHijo("IDUdMedida")
    '                drT("NoAcumular") = dblAcumular
    '                drT("Nivel") = Nivel
    '                drT("TipoLinea") = enumTipoLineaTrabajo.tltTrabajo
    '                If drT("TipoLinea") = enumTipoLineaTrabajo.tltTrabajo And InStr(drTrabajoHijo("CODIGOHIJO"), "@") = 0 Then
    '                    If Length(drTrabajoHijo("TotalMedicion")) > 0 Then
    '                        drT("QPresup") = drTrabajoHijo("TotalMedicion")
    '                    Else
    '                        drT("QPresup") = 0
    '                    End If
    '                    drT("IDTrabajoPresupCopia") = IDTrabajo
    '                    drT("ImpPresupTrabajoA") = drTrabajoHijo("Precio")
    '                    drT("ImpPresupTrabajoVentaA") = drTrabajoHijo("Precio")
    '                    drT("ImpPresupQTrabajoA") = xRound(drTrabajoHijo("Precio") * drT("QPresup"), 2)
    '                    drT("ImpPresupQTrabajoVentaA") = xRound(drTrabajoHijo("Precio") * drT("QPresup"), 2)
    '                Else
    '                    drT("QPresup") = 0
    '                End If
    '                drT("IDTrabajoPresupCopia") = IDTrabajo
    '                drT("IdTrabajoPresupPadre") = IDTrabajo
    '                IDTrabajoHijo = drT("IDTrabajoPresup")
    '                drT("Secuencia") = Orden
    '                Orden = Orden + 10
    '                If CodTrabajoHijo <> "" Then
    '                    CodTrabajoHijorecursivo = CodTrabajoHijo
    '                    CodTrabajoNietorecursivo = drTrabajoHijo("CodTrabajo")
    '                End If

    '                If CodTrabajoHijo = "" Then
    '                    CodTrabajoHijorecursivo = drTrabajoHijo("CodTrabajo")
    '                End If

    '                dtTrabajosN.Rows.Add(drT.ItemArray)

    '                If blnImportarMateriales Then CrearMateriales(NumPresup, drTrabajoHijo("CodTrabajo"), drTrabajoHijo("IDTrabajoPresup"), dtMaterialesN, Nz(drTrabajoHijo("TotalMedicion"), 0))
    '                'If blnImportarMOD Then CrearManoObra(NumPresup, drTrabajoHijo("CodTrabajo"), strTipoHora, drTrabajoHijo("IDTrabajoPresup"), dtManoObraN, Nz(drTrabajoHijo("TotalMedicion"), 0))
    '                'If blnImportarCentros Then CrearCentros(NumPresup, drTrabajoHijo("CodTrabajo"), drTrabajoHijo("IDTrabajoPresup"), dtCentroTrabajoN, Nz(drTrabajoHijo("TotalMedicion"), 0))
    '                'If blnImportarMediciones Then CrearMediciones(NumPresup, drTrabajoHijo("CodTrabajo"), drTrabajoHijo("IDTrabajoPresup"), dtMedicionesN)
    '            End If

    '            Dim strSql As String = "DELETE FROM FORMATO_D WHERE ID=" & drTrabajoHijo("IDTrabajoPresup")
    '            AdminData.Execute(strSql)

    '            CrearNivelesRecursivos(NumPresup, dblAcumular, Nivel, Orden, IDTrabajoHijo, strIDArticulo, strTipoHora, dtTrabajosN, _
    '                                   dtObraSubTipoTrabajoN, dtObraSubSubTipoTrabajoN, dtMaterialesN, dtManoObraN, dtCentroTrabajoN, dtMedicionesN, _
    '                                   blnImportarMateriales, blnImportarMOD, blnImportarCentros, blnImportarMediciones, _
    '                                   CodigoPadre, CodTrabajo, CodTrabajoHijorecursivo, CodTrabajoNietorecursivo)
    '        Next
    '    End If
    'End Sub
    Public Sub CrearNivelesFormatoC(ByRef data As DataTrabajo)
        'Trabajos Dependientes

        Dim ObraSubTipoTrabajo As New ObraSubtipoTrabajo
        Dim drObraSubTipoTrabajo As DataRow
        Dim ObraSubSubTipoTrabajo As New ObraSubSubtipoTrabajo
        Dim drObraSubSubTipoTrabajo As DataRow
        Dim drT As DataRow
        Dim dtExiste As DataTable
        Dim dtTrabajos As DataTable
        Dim IDTrabajoHijo As Integer
        Dim CodTrabajoHijorecursivo As String = String.Empty
        Dim CodTrabajoNietorecursivo As String = String.Empty
        Dim IDSubSubTipoTrabajo As String
        Dim IDSubTipoTrabajo As String
        Dim IDSubSubTipoTrabajoC As String
        Dim IDSubTipoTrabajoC As String
        Dim IDTipoObra As String
        Dim IDTipoTrabajo As String
        Dim Rama As Integer
        Dim Nivel As Integer



        dtTrabajos = New BE.DataEngine().Filter("vBC3CapitulosHijos", "", "CodTrabajo not like '%@@%'", "ID")

        If Not dtTrabajos Is Nothing AndAlso dtTrabajos.Rows.Count > 0 Then

            For Each drTrabajoHijo As DataRow In dtTrabajos.Select
                drT = data.dtTrabajosN.NewRow
                drT("IDTrabajoPresup") = drTrabajoHijo("IDHijo")
                drT("IDPresup") = data.NumPresup
                drT("IDTipoObra") = data.IDTipoObra
                drT("CodTrabajo") = drTrabajoHijo("CodigoHijo")
                drT("DescTrabajo") = drTrabajoHijo("DescTrabajo")
                drT("IDArticulo") = data.strIDArticulo
                drT("IDUdMedida") = drTrabajoHijo("IDUdMedida")
                drT("NoAcumular") = data.dblAcumular
                drT("Nivel") = data.nivel


              
                If Length(drTrabajoHijo("TotalMedicion")) > 0 Then
                    drT("QPresup") = drTrabajoHijo("TotalMedicion")
                Else
                    drT("QPresup") = 0
                End If
                drT("QUnidad") = drT("QPresup")

                Dim oFilter As New Filter
                oFilter.Add(New NumberFilterItem("IDTrabajoPresup", FilterOperator.Equal, drTrabajoHijo("IDPadre")))
                Dim dvTrabajosN As DataView = data.dtTrabajosN.DefaultView
                dvTrabajosN.RowFilter = oFilter.Compose(New AdoFilterComposer)
                If dvTrabajosN.Count > 0 Then
                    'Si el padre era un capítulo acumula
                    If Length(drTrabajoHijo("CODIGOHIJOINICIAL")) > 0 AndAlso InStr(drTrabajoHijo("CODIGOHIJOINICIAL"), "@") Then
                        drT("TipoLinea") = enumTipoLineaTrabajo.tltCapitulo
                    Else
                        If dvTrabajosN.Item(0)("TipoLinea") = enumTipoLineaTrabajo.tltCapitulo Then
                            If drT("QPresup") = 0 Then drT("QPresup") = 1
                            drT("TipoLinea") = enumTipoLineaTrabajo.tltTrabajo
                        Else
                            drT("TipoLinea") = enumTipoLineaTrabajo.tltTrabajoConcepto
                        End If
                    End If
                  


                    drT("QPresup") = drT("QPresup") * Nz(dvTrabajosN.Item(0)("QPresup"), 1)
                    If Length(dvTrabajosN.Item(0)("IDTipoTrabajo")) > 0 Then
                        drT("IDTipoTrabajo") = dvTrabajosN.Item(0)("IDTipoTrabajo")
                    End If
                    If Length(dvTrabajosN.Item(0)("IDSubTipoTrabajo")) > 0 Then
                        drT("IDSubTipoTrabajo") = dvTrabajosN.Item(0)("IDSubTipoTrabajo")
                        drT("Nivel") = drT("Nivel") + 1
                    Else
                        'Crear subtipo

                        Dim f As New Filter
                        f.Add(New StringFilterItem("IdTipoObra", data.IDTipoObra))
                        f.Add(New StringFilterItem("IDTipoTrabajo", drT("IDTipoTrabajo")))
                        f.Add(New StringFilterItem("IDSubTipoTrabajo", drTrabajoHijo("CodigoHijo")))
                        dtExiste = ObraSubTipoTrabajo.Filter(f)
                        If dtExiste.Rows.Count = 0 Then
                            Dim dtObrasubTipo As DataTable = ObraSubTipoTrabajo.AddNew
                            drObraSubTipoTrabajo = data.dtObraSubTipoTrabajoN.NewRow
                            drObraSubTipoTrabajo("IDTipoObra") = data.IDTipoObra
                            drObraSubTipoTrabajo("IDTipoTrabajo") = drT("IDTipoTrabajo")
                            drObraSubTipoTrabajo("IDSubTipoTrabajo") = drTrabajoHijo("CodigoHijo")
                            drObraSubTipoTrabajo("DescsubtipoTrabajo") = drTrabajoHijo("DescTrabajo")
                            dtObrasubTipo.Rows.Add(drObraSubTipoTrabajo.ItemArray)
                            If Length(drObraSubTipoTrabajo("IDTipoTrabajo")) > 0 Then
                                BusinessHelper.UpdatePackage(New UpdatePackage(dtObrasubTipo.TableName, dtObrasubTipo))
                            End If
                        End If
                        drT("IDSubTipoTrabajo") = drTrabajoHijo("CodigoHijo")
                    End If
                Else
                    drT("TipoLinea") = enumTipoLineaTrabajo.tltTrabajo
                End If

                drT("IDTrabajoPresupCopia") = data.IDHijo
                drT("ImpPresupTrabajoA") = drTrabajoHijo("Precio")
                drT("ImpPresupTrabajoVentaA") = drTrabajoHijo("Precio")
                drT("ImpPresupQTrabajoA") = xRound(drTrabajoHijo("Precio") * drT("QPresup"), 2)
                drT("ImpPresupQTrabajoVentaA") = xRound(drTrabajoHijo("Precio") * drT("QPresup"), 2)




                drT("IDTrabajoPresupCopia") = drTrabajoHijo("IDHijo")
                drT("IdTrabajoPresupPadre") = drTrabajoHijo("IDPadre")

                IDTrabajoHijo = drT("IDTrabajoPresup")

                drT("Secuencia") = data.orden
                data.orden = data.orden + 10

                data.dtTrabajosN.Rows.Add(drT.ItemArray)
                data.cantidad = Nz(drT("QPresup"), 0)
                data.IDPadre = drTrabajoHijo("IDHijo")
                Dim dtFormato_C As DataTable = New BE.DataEngine().Filter("vBC3CodigosIniciales", New StringFilterItem("CODIGO", drTrabajoHijo("CodigoHijo")))
                If Not dtFormato_C Is Nothing AndAlso dtFormato_C.Rows.Count > 0 Then
                    data.IDFiltro = dtFormato_C.Rows(0)("ID")
                End If
                If data.blnImportarMateriales Then CrearMateriales(data)
                If data.blnImportarMOD Then CrearManoObra(data)
                If data.blnImportarCentros Then CrearCentros(data)
                If data.blnImportarMediciones Then CrearMediciones(data)


                Dim strSql As String = "DELETE FROM FORMATO_D WHERE ID=" & drTrabajoHijo("ID")
                AdminData.Execute(strSql)

            Next
        End If
    End Sub
 
    Public Sub CrearPresupuesto(ByVal NumPresup As Integer, _
                                Optional ByVal blnImportarMateriales As Boolean = True, _
                                Optional ByVal blnCrearMateriales As Boolean = True, _
                                Optional ByVal blnImportarMOD As Boolean = True, _
                                Optional ByVal blnImportarCentros As Boolean = True, _
                                Optional ByVal blnImportarMediciones As Boolean = True)

        AdminData.Execute("UPDATE FORMATO_M SET FORMULA = COMENTARIO WHERE TIPO = 3")

        Dim CodigoPadre As String = CrearTipoPresupuesto(NumPresup)
        If blnImportarMateriales Then
            CrearUnidadesMedida(False)
        Else
            CrearUnidadesMedida(True)
        End If
        CrearTrabajos(NumPresup, CodigoPadre, blnImportarMateriales, blnCrearMateriales, blnImportarMOD, blnImportarCentros, blnImportarMediciones)

        Dim sqltexto As String
        sqltexto = "update tbObraTrabajoPresup set Texto=TextoLargo from " _
        & "(select IdTrabajoPresup, Texto as TextoLargo from vObraDatosTextoEntero) otpp " _
        & " inner join tbObraTrabajoPresup otp on otpp.IDTrabajoPresup = otp.IDTrabajoPresup where IDpresup=" & NumPresup
        AdminData.Execute(sqltexto)
        
        sqltexto = "update tbObraTrabajoPresup set IDTrabajoPresupCopia = null"
        AdminData.Execute(sqltexto)

        sqltexto = "UPDATE tbObraTrabajoPresup SET TipoLinea = 4 FROM  tbObraTrabajoPresup INNER JOIN  vObraTrabajoPresupPadre ON tbObraTrabajoPresup.IDTrabajoPresupPadre = vObraTrabajoPresupPadre.IDTrabajoPresup WHERE  (vObraTrabajoPresupPadre.TipoPadre = 0) and IDpresup=" & NumPresup
        AdminData.Execute(sqltexto)

        Dim services As New ServiceProvider
        ProcessServer.ExecuteTask(Of Integer)(AddressOf ObraPresupCabecera.RecalcularPresupuesto, NumPresup, services)
        ProcessServer.ExecuteTask(Of Integer)(AddressOf RecalcularSecuenciaTrabajo, NumPresup, services)
    End Sub

    <Task()> Public Shared Sub RecalcularSecuenciaTrabajo(ByVal IDPresup As Integer, ByVal services As ServiceProvider)
        Dim dtTrabajos As DataTable = New ObraTrabajoPresup().Filter(New NumberFilterItem("IDPresup", IDPresup), "IDTrabajoPresup")
        If dtTrabajos.Rows.Count > 0 Then
            Dim where As String = New IsNullFilterItem("IDTrabajoPresupPadre", True).Compose(New AdoFilterComposer)
            Dim d As New dataSecuencia(10)
            For Each drTrabajo As DataRow In dtTrabajos.Select(where, "Secuencia")
                drTrabajo("Secuencia") = d.Secuencia
                d.Secuencia += 10
                d.drTrabajo = drTrabajo
                d.dtTrabajos = dtTrabajos
                ProcessServer.ExecuteTask(Of dataSecuencia)(AddressOf RecalcularSecuenciaTrabajoHijos, d, services)
            Next
            ObraTrabajoPresup.UpdateTable(dtTrabajos)
        End If
    End Sub

    Public Class dataSecuencia
        Public drTrabajo As DataRow
        Public dtTrabajos As DataTable
        Public Secuencia As Integer

        Public Sub New(ByVal Secuencia As Integer)
            Me.Secuencia = Secuencia
        End Sub
    End Class
    <Task()> Public Shared Sub RecalcularSecuenciaTrabajoHijos(ByVal data As dataSecuencia, ByVal services As ServiceProvider)
        If Length(data.drTrabajo("IDTrabajoPresup")) > 0 Then
            Dim dv As DataView = data.dtTrabajos.DefaultView
            Dim where As String = New StringFilterItem("IDTrabajoPresupPadre", data.drTrabajo("IDTrabajoPresup")).Compose(New AdoFilterComposer)

            For Each drTrabajo As DataRow In data.dtTrabajos.Select(where, "Secuencia")
                drTrabajo("Secuencia") = data.Secuencia
                data.Secuencia += 10
                data.drTrabajo = drTrabajo
                ProcessServer.ExecuteTask(Of dataSecuencia)(AddressOf RecalcularSecuenciaTrabajoHijos, data, services)
            Next
        End If
    End Sub

End Class
