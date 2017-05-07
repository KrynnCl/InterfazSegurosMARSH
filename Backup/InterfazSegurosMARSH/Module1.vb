Imports System.Net
Imports System.IO
Imports InterfazSegurosMARSH.Codigos


Module Module1


    Public Sub Main(ByVal args() As String)

        '**************************
        'Carga de Los contratos
        'Mientras se piensa pro distrito
        '**************************
        CargaOC_COntratos("CMLB")
        'CargaOC_COntratos("ALTO")
        '**************************

        '**************************
        'Marca todos los contratos con alguna versión mayor que no se ha asegurado
        Dim TablaSeguros As New SegurosMarshTableAdapters.SegurosMARSH_SeguroTableAdapter
        TablaSeguros.InsertarInoperantes()
        '***************************************

        '**************************
        'Procedimiento para Crear los Seguros correspondientes.
        CrearSeguros()
        '***************************************

        '**************************
        'Procedimiento que envia los avisos correspondientes

        Send_Avisos_Seguros_OC()
        '***************************************

    End Sub 'Main

    Public Sub CrearSeguros()
        '**************************
        'Obtencion de los XML Para crear nuevos Seguros
        '**************************
        Dim TablaContratos As SegurosMarshTableAdapters.SegurosMARSH_Oc_ContratoTableAdapter
        Dim tContratos As SegurosMarsh.SegurosMARSH_Oc_ContratoDataTable
        Dim fContrato As SegurosMarsh.SegurosMARSH_Oc_ContratoRow
        TablaContratos = New SegurosMarshTableAdapters.SegurosMARSH_Oc_ContratoTableAdapter
        Dim creador As New System.Xml.XmlDocument()
        Dim strXML, strXMLResponse As String
        Dim TablaSeguros As New SegurosMarshTableAdapters.SegurosMARSH_SeguroTableAdapter
        Dim MontoDeclarado, Prima As Double

        '****************************
        'Creacion de ORdenes de Compra!!!!!!!!!!!
        tContratos = TablaContratos.GetOCSinSeguro

        If tContratos.Count > 0 Then

            For Each fContrato In tContratos
                strXML = Get_XML_CONTRATO(fContrato.Cod_contrat, fContrato.Distrito, fContrato.Version, fContrato.TipoContr)

                strXMLResponse = CargaMARSH(strXML)


                creador.LoadXml(strXMLResponse)
                'Ingreso de la REspuesta en la BD
                If creador.GetElementsByTagName("Prima_contrat").Item(0).InnerText = "" Then
                    Prima = 0
                Else
                    Prima = CDbl(creador.GetElementsByTagName("Prima_contrat").Item(0).InnerText)
                End If
                If creador.GetElementsByTagName("Mont_total_asegurado").Item(0).InnerText = "" Then
                    MontoDeclarado = 0
                Else
                    MontoDeclarado = CDbl(creador.GetElementsByTagName("Mont_total_asegurado").Item(0).InnerText)
                End If
                Dim VersionSeguro As String
                VersionSeguro = GetVersionSeguro(fContrato.Cod_contrat, fContrato.Distrito, fContrato.Version, fContrato.TipoContr)
                If VersionSeguro = "" Then
                    VersionSeguro = "00"
                End If
                VersionSeguro = Strings.Right(VersionSeguro, 2)
                TablaSeguros.Insert(fContrato.Cod_contrat, _
                                    fContrato.Distrito, fContrato.TipoContr, fContrato.Version, _
                                    creador.GetElementsByTagName("Id_verificacion").Item(0).InnerText, _
                                    creador.GetElementsByTagName("F_inicio_declaracion").Item(0).InnerText, _
                                    creador.GetElementsByTagName("F_term_declaracion").Item(0).InnerText, _
                                    Prima, _
                                    MontoDeclarado, _
                                    creador.GetElementsByTagName("Aprobado").Item(0).InnerText, _
                                    creador.GetElementsByTagName("Motivo").Item(0).InnerText, _
                                    creador.GetElementsByTagName("URL_Certificado").Item(0).InnerText, fContrato.Correo_aviso, "", Format(Today, "yyyy-MM-dd"), VersionSeguro)

            Next
        End If


        '****************************
        'Creacion de Contratos!!!!!!!!!!!
        tContratos = TablaContratos.GetContratosSinSeguro
        If tContratos.Count > 0 Then

            For Each fContrato In tContratos
                strXML = Get_XML_CONTRATO(fContrato.Cod_contrat, fContrato.Distrito, fContrato.Version, fContrato.TipoContr)

                strXMLResponse = CargaMARSH(strXML)


                creador.LoadXml(strXMLResponse)
                'Ingreso de la REspuesta en la BD
                If creador.GetElementsByTagName("Prima_contrat").Item(0).InnerText = "" Then
                    Prima = 0
                Else
                    Prima = CDbl(creador.GetElementsByTagName("Prima_contrat").Item(0).InnerText)
                End If
                If creador.GetElementsByTagName("Mont_total_asegurado").Item(0).InnerText = "" Then
                    MontoDeclarado = 0
                Else
                    MontoDeclarado = CDbl(creador.GetElementsByTagName("Mont_total_asegurado").Item(0).InnerText)
                End If
                Dim VersionSeguro As String
                VersionSeguro = GetVersionSeguro(fContrato.Cod_contrat, fContrato.Distrito, fContrato.Version, fContrato.TipoContr)
                If VersionSeguro = "" Then
                    VersionSeguro = "00"
                End If
                VersionSeguro = Strings.Right(VersionSeguro, 2)

                TablaSeguros.Insert(fContrato.Cod_contrat, _
                                    fContrato.Distrito, fContrato.TipoContr, fContrato.Version, _
                                    creador.GetElementsByTagName("Id_verificacion").Item(0).InnerText, _
                                    creador.GetElementsByTagName("F_inicio_declaracion").Item(0).InnerText, _
                                    creador.GetElementsByTagName("F_term_declaracion").Item(0).InnerText, _
                                    Prima, _
                                    MontoDeclarado, _
                                    creador.GetElementsByTagName("Aprobado").Item(0).InnerText, _
                                    creador.GetElementsByTagName("Motivo").Item(0).InnerText, _
                                    creador.GetElementsByTagName("URL_Certificado").Item(0).InnerText, fContrato.Correo_aviso, "", Format(Today, "yyyy-MM-dd"), VersionSeguro)

            Next
        End If

    End Sub
    'Procedimiento para Avisar ingresos
    Public Sub Send_Avisos_Seguros()
        Dim Mensage, MensageOC, MensageContrato, MensageAux As String
        Dim taSegurosSinAviso As New SegurosMarshTableAdapters.SegurosMARSH_SeguroTableAdapter
        Dim tSeguros As SegurosMarsh.SegurosMARSH_SeguroDataTable
        tSeguros = taSegurosSinAviso.GetSegurosPorCorreo
        ' tSeguros = taSegurosSinAviso.GetDataSeguro
        Dim tseguro As SegurosMarsh.SegurosMARSH_SeguroRow
        Dim Email As New ArrayList


        'Dim eeemail As Array
        'Email.Resize(Email, 3)
        Dim TablaConfiguracion As New SegurosMarshTableAdapters.SegurosMARSH_ConfiguracionTableAdapter()
        'Email.AddRange(TablaConfiguracion.GetCorreos("Admin").GetEnumerator)
        Dim x As Integer = 1
        For Each Correo As SegurosMarsh.SegurosMARSH_ConfiguracionRow In TablaConfiguracion.GetCorreos("Admin")
            Email.Add(Correo.Valor2)
            x += 1
        Next
        Email.Insert(0, "")

        Mensage = " Estimado, en la ultima actualización de seguros MARSH se cuentan los siguientes cambios: <br><br>"
        If tSeguros.Count = 0 Then
            MensageOC = "No se han encontrado seguros asociados a ordenes o contratos<br><br>"

        Else
            'Ordenes de compra
            MensageOC = "Se han ingresado las siguientes ordenes de compra: <br><br>"
            MensageOC += "<table width='100%' border = '1' > "
            MensageOC += "<tr><td><b>OC</b></td>"
            MensageOC += "<td><b>Versión Seguro</b></td> "
            MensageOC += "<td><b>Id Verificacion</b></td> "
            MensageOC += "<td><b>Aprobado</b></td>"
            MensageOC += "<td><b>Motivo</b></td>"
            MensageOC += "<td><b>URL Certificado</b></td>"
            MensageOC += "</tr>"

            MensageContrato = "<br><br>Se han ingresado los siguientes Contratos: <br><br>"
            MensageContrato += "<table width='100%' border = '1' > "
            MensageContrato += "<tr><td><b>Contrato</b></td>"
            MensageContrato += "<td><b>Versión Seguro</b></td> "
            MensageContrato += "<td><b>Id Verificacion</b></td> "
            MensageContrato += "<td><b>Aprobado</b></td>"
            MensageContrato += "<td><b>Motivo</b></td>"
            MensageContrato += "<td><b>URL Certificado</b></td>"
            MensageContrato += "</tr>"



            For Each tseguro In tSeguros


                If Email(0) = "" Then
                    Email(0) = tseguro.EmailAviso
                End If

                If Email(0) <> tseguro.EmailAviso Then

                    Mensage += MensageOC
                    Mensage += "</Table>"
                    Mensage += MensageContrato
                    Mensage += "</Table>"


                    SendMail(Email, Mensage, "Interfaz MARSH:: Creación de Nuevos Seguros")
                    'taSegurosSinAviso.ActualizarEnviados()
                    Email(0) = tseguro.EmailAviso
                    Mensage = " Estimado, en la ultima actualización de seguros MARSH se cuentan los siguientes cambios: <br><br>"

                    MensageOC = "Se han ingresado las siguientes ordenes de compra: <br><br>"
                    MensageOC += "<table width='100%' border = '1' > "
                    MensageOC += "<tr><td><b>OC</b></td>"
                    MensageOC += "<td><b>Versión Seguro</b></td> "
                    MensageOC += "<td><b>Id Verificacion</b></td> "
                    MensageOC += "<td><b>Aprobado</b></td>"
                    MensageOC += "<td><b>Motivo</b></td>"
                    MensageOC += "<td><b>URL Certificado</b></td>"
                    MensageOC += "</tr>"

                    MensageContrato = "<br><br>Se han ingresado los siguientes Contratos: <br><br>"
                    MensageContrato += "<table width='100%' border = '1' > "
                    MensageContrato += "<tr><td><b>Contrato</b></td>"
                    MensageContrato += "<td><b>Versión Seguro</b></td> "
                    MensageContrato += "<td><b>Id Verificacion</b></td> "
                    MensageContrato += "<td><b>Aprobado</b></td>"
                    MensageContrato += "<td><b>Motivo</b></td>"
                    MensageContrato += "<td><b>URL Certificado</b></td>"
                    MensageContrato += "</tr>"
                End If
                MensageAux = "<tr><td><b>" & tseguro.Cod_contrat & " </b></td>"
                MensageAux += "<td><b>" & tseguro.Version_Seguro & "</b></td> "
                MensageAux += "<td><b>" & tseguro.Id_verificacion & "</b></td> "
                MensageAux += "<td><b>" & tseguro.Aprobado & "</b></td>"
                MensageAux += "<td><b>" & tseguro.Motivo & "</b></td>"
                MensageAux += "<td><b>" & tseguro.URL_Certificado & "</b></td>"
                MensageAux += "</tr>"

                If tseguro.TipoContr = "C" Then
                    MensageContrato += MensageAux
                Else
                    MensageOC += MensageAux
                End If

            Next
            Mensage += MensageOC
            Mensage += "</Table>"
            Mensage += MensageContrato
            Mensage += "</Table>"
            SendMail(Email, Mensage, "Interfaz MARSH:: Creación de Nuevos Seguros")
            taSegurosSinAviso.ActualizarEnviados()


            ''Contratos



        End If

    End Sub

    Public Sub Send_Avisos_Seguros_OC()
        Dim Mensage, MensageOC, MensageContrato, MensageAux As String
        Dim taSegurosSinAviso As New SegurosMarshTableAdapters.SegurosMARSH_SeguroTableAdapter
        Dim tSeguros As SegurosMarsh.SegurosMARSH_SeguroDataTable
        tSeguros = taSegurosSinAviso.GetSegurosPorCorreo
        ' tSeguros = taSegurosSinAviso.GetDataSeguro
        Dim tseguro As SegurosMarsh.SegurosMARSH_SeguroRow
        Dim Email As New ArrayList


        'Dim eeemail As Array
        'Email.Resize(Email, 3)
        Dim TablaConfiguracion As New SegurosMarshTableAdapters.SegurosMARSH_ConfiguracionTableAdapter()
        'Email.AddRange(TablaConfiguracion.GetCorreos("Admin").GetEnumerator)
        Dim x As Integer = 1
        For Each Correo As SegurosMarsh.SegurosMARSH_ConfiguracionRow In TablaConfiguracion.GetCorreos("Admin")
            Email.Add(Correo.Valor2)
            x += 1
        Next
        Email.Insert(0, "")

        Mensage = " Estimado, en la ultima actualización de seguros MARSH se cuentan los siguientes cambios: <br><br>"
        If tSeguros.Count = 0 Then
            MensageOC = "No se han encontrado seguros asociados a ordenes o contratos<br><br>"

        Else
            'Ordenes de compra
            MensageOC = "Se han ingresado las siguientes ordenes de compra: <br><br>"
            MensageOC += "<table width='100%' border = '1' > "
            MensageOC += "<tr><td><b>OC</b></td>"
            MensageOC += "<td><b>Versión Seguro</b></td> "
            MensageOC += "<td><b>Id Verificacion</b></td> "
            MensageOC += "<td><b>Aprobado</b></td>"
            MensageOC += "<td><b>Motivo</b></td>"
            MensageOC += "<td><b>URL Certificado</b></td>"
            MensageOC += "</tr>"

          



            For Each tseguro In tSeguros


                If Email(0) = "" Then
                    Email(0) = tseguro.EmailAviso
                End If

                If Email(0) <> tseguro.EmailAviso Then

                    Mensage += MensageOC
                    Mensage += "</Table>"
              

                    SendMail(Email, Mensage, "Interfaz MARSH:: Creación de Nuevos Seguros Asociado a OC")
                    'taSegurosSinAviso.ActualizarEnviados()
                    Email(0) = tseguro.EmailAviso
                    Mensage = " Estimado, en la ultima actualización de seguros MARSH se cuentan los siguientes cambios: <br><br>"

                    MensageOC = "Se han ingresado las siguientes ordenes de compra: <br><br>"
                    MensageOC += "<table width='100%' border = '1' > "
                    MensageOC += "<tr><td><b>OC</b></td>"
                    MensageOC += "<td><b>Versión Seguro</b></td> "
                    MensageOC += "<td><b>Id Verificacion</b></td> "
                    MensageOC += "<td><b>Aprobado</b></td>"
                    MensageOC += "<td><b>Motivo</b></td>"
                    MensageOC += "<td><b>URL Certificado</b></td>"
                    MensageOC += "</tr>"

                  
                End If
                MensageAux = "<tr><td><b>" & tseguro.Cod_contrat & " </b></td>"
                MensageAux += "<td><b>" & tseguro.Version_Seguro & "</b></td> "
                MensageAux += "<td><b>" & tseguro.Id_verificacion & "</b></td> "
                MensageAux += "<td><b>" & tseguro.Aprobado & "</b></td>"
                MensageAux += "<td><b>" & tseguro.Motivo & "</b></td>"
                MensageAux += "<td><b>" & tseguro.URL_Certificado & "</b></td>"
                MensageAux += "</tr>"

                If tseguro.TipoContr = "C" Then
                    '        MensageContrato += MensageAux
                Else
                    MensageOC += MensageAux
                End If

            Next
            Mensage += MensageOC
            Mensage += "</Table>"
            
            SendMail(Email, Mensage, "Interfaz MARSH:: Creación de Nuevos Seguros Asociado a OC")
            taSegurosSinAviso.ActualizarEnviados()


            ''Contratos



        End If

    End Sub


    'Procedimiento que Entrega el XML de una fila identificada.
    Public Function Get_XML_CONTRATO(ByVal Contra_OC As String, ByVal pDistrito As String, ByVal pVersion As String, ByVal pTipo As String) As String
        Dim TablaContratos As SegurosMarshTableAdapters.SegurosMARSH_Oc_ContratoTableAdapter
       
        TablaContratos = New SegurosMarshTableAdapters.SegurosMARSH_Oc_ContratoTableAdapter

        Dim VersionSeguro As String
        VersionSeguro = GetVersionSeguro(Contra_OC, pDistrito, pVersion, pTipo)

        Return (TablaContratos.Get_XML_CONTRATO_OC(VersionSeguro, Contra_OC, pDistrito, pTipo, pVersion))

    End Function

    Public Function GetVersionSeguro(ByVal Contra_OC As String, ByVal pDistrito As String, ByVal pVersion As String, ByVal pTipo As String) As String
        Dim VersionSeguro As String
        Dim iVersion As Integer
        Dim TablaSeguros As New SegurosMarshTableAdapters.SegurosMARSH_SeguroTableAdapter
        VersionSeguro = TablaSeguros.GetUltimaVersionSeguro(pDistrito, Contra_OC, pTipo)

        If VersionSeguro <> Nothing Then
            iVersion = CInt(VersionSeguro)
            iVersion = iVersion + 1
            VersionSeguro = "0" & CStr(iVersion)
            VersionSeguro = Strings.Right(VersionSeguro, 2)
            VersionSeguro = "-" & VersionSeguro
        Else
            VersionSeguro = ""
        End If
        GetVersionSeguro = VersionSeguro
    End Function
    'Procedimiento donde se cargarán todos los Contratos y Ordenes de Compra de un Distrito especifico para el
    'momento en que se ejecuta la interfaz.
    'Datos desde ELLPR -> BD Interfaz
    Public Sub CargaOC_COntratos(ByVal pDistrito As String)
        'Contadores.
        Dim cOrdenes_Nuevas, cOrdenes_Version_nuevas, cContratos_Nuevos, cContratos_Version_Nuevos As Integer
        Dim Fecha_Termino As String
        '
        Dim dtOC_Ellipse As New SegurosMarshTableAdapters.Ellipse_OCTableAdapter
        Dim fOC_Ellipse As SegurosMarsh.Ellipse_OCRow
        Dim dOC_Ellipse As SegurosMarsh.Ellipse_OCDataTable

        Dim dtSMarsh As New SegurosMarshTableAdapters.SegurosMARSH_Oc_ContratoTableAdapter

        '******************** ORDENES DE COMPRAAAAAAAAA*********************************
        dOC_Ellipse = dtOC_Ellipse.GetOrdenes_compra

        For Each fOC_Ellipse In dOC_Ellipse

            Dim MontoContrato As Double
            'Monto Minimo de Contrato es de 2 UF 
            '.- 20110920 .- Modificado Monto Contrato = 20
            MontoContrato = 2
            If fOC_Ellipse.MONT_CONTRAT < MontoContrato Then
                If fOC_Ellipse.MONT_ORDEN_UF >= MontoContrato Then
                    MontoContrato = fOC_Ellipse.MONT_ORDEN_UF
                End If
            Else
                MontoContrato = fOC_Ellipse.MONT_CONTRAT
            End If
            Dim DAnual As String

            'If DateDiff(DateInterval.Month, CDate(fOC_Ellipse.F_INICIO_VIG), CDate(fOC_Ellipse.F_TERM_VIG)) > 36 Then
            '    DAnual = "Si"
            'Else
            '    DAnual = "No"

            'End If
            '************* Esto cambia por correo de Carlos donde TODAS las Ordenes son DANUAL NO.
            '.-20110920 Cambio de DANUAL SIEMPRE A NO
            DAnual = "No"

            '.-20110920 - Se agrega cambio de Fecha de termino es fecha de termino de la OC mas 60 dias
            Fecha_Termino = Format(DateAdd(DateInterval.Day, 60, CDate(fOC_Ellipse.F_TERM_VIG)), "yyyy-MM-dd")

            If (dtSMarsh.Existe_OC_Contrato(fOC_Ellipse.DISTRITO, "O", fOC_Ellipse.COD_CONTRAT) = 0) Then

                dtSMarsh.Insert(fOC_Ellipse.EMPRESA, fOC_Ellipse.RESP_CONTRAT, fOC_Ellipse.D_ANUAL, fOC_Ellipse.F_INICIO_VIG _
                                , Fecha_Termino, fOC_Ellipse.N_POLIZA, fOC_Ellipse.RUT_CONT, fOC_Ellipse.DIG_RUT_CONT _
                                , fOC_Ellipse.R_SOC_CONT, fOC_Ellipse.DIR_CONT, fOC_Ellipse.CIUD_CONT, fOC_Ellipse.NOMB_CONTACT _
                                , fOC_Ellipse.COD_CONTRAT, MontoContrato, fOC_Ellipse.TIP_CONTRAT, fOC_Ellipse.DESCRIP_TRAB _
                                , fOC_Ellipse.DISTRITO, "O", "00", Format(Today, "yyyy-MM-dd"), fOC_Ellipse.EMAIL_RESPONSABLE_CONT)
                cOrdenes_Nuevas += 1


            Else
                If (dtSMarsh.Existe_Identico(fOC_Ellipse.DISTRITO, fOC_Ellipse.COD_CONTRAT, fOC_Ellipse.D_ANUAL _
                                             , fOC_Ellipse.F_INICIO_VIG, Fecha_Termino, fOC_Ellipse.N_POLIZA, fOC_Ellipse.RUT_CONT _
                                             , fOC_Ellipse.DIG_RUT_CONT, MontoContrato, fOC_Ellipse.DESCRIP_TRAB, fOC_Ellipse.EMPRESA _
                                             , fOC_Ellipse.RESP_CONTRAT, fOC_Ellipse.DIR_CONT, fOC_Ellipse.R_SOC_CONT, fOC_Ellipse.CIUD_CONT, fOC_Ellipse.NOMB_CONTACT, fOC_Ellipse.TIP_CONTRAT, "O") = 0) Then
                    Dim Version As String
                    Version = dtSMarsh.GetUltimaVersion(fOC_Ellipse.DISTRITO, fOC_Ellipse.COD_CONTRAT, "O")
                    Dim iVersion As Integer
                    iVersion = CInt(Version)
                    iVersion = iVersion + 1
                    Version = "0" & CStr(iVersion)
                    Version = Strings.Right(Version, 2)
                    dtSMarsh.Insert(fOC_Ellipse.EMPRESA, fOC_Ellipse.RESP_CONTRAT, fOC_Ellipse.D_ANUAL, fOC_Ellipse.F_INICIO_VIG _
                                , Fecha_Termino, fOC_Ellipse.N_POLIZA, fOC_Ellipse.RUT_CONT, fOC_Ellipse.DIG_RUT_CONT _
                                , fOC_Ellipse.R_SOC_CONT, fOC_Ellipse.DIR_CONT, fOC_Ellipse.CIUD_CONT, fOC_Ellipse.NOMB_CONTACT _
                                , fOC_Ellipse.COD_CONTRAT, MontoContrato, fOC_Ellipse.TIP_CONTRAT, fOC_Ellipse.DESCRIP_TRAB _
                                , fOC_Ellipse.DISTRITO, "O", Version, Format(Today, "yyyy-MM-dd"), fOC_Ellipse.EMAIL_RESPONSABLE_CONT)
                    cOrdenes_Version_nuevas += 1
                    'En cualquier otro caso, no existen cambios

                End If
            End If
        Next
        '******************** FIN   ORDENES DE COMPRAAAAAAAAA*********************************


        '******************** CONTRATOS+++++++++++++++++++*********************************

        'Dim dtCont_Ellipse As New SegurosMarshTableAdapters.Ellipse_ContratosTableAdapter
        'Dim fCont_Ellipse As SegurosMarsh.Ellipse_ContratosRow
        'Dim dCont_Ellipse As SegurosMarsh.Ellipse_ContratosDataTable


        'dCont_Ellipse = dtCont_Ellipse.GetContratos
        'For Each fCont_Ellipse In dCont_Ellipse


        '    If (dtSMarsh.Existe_OC_Contrato(fCont_Ellipse.DISTRITO, "C", fCont_Ellipse.COD_CONTRAT) = 0) Then
        '        dtSMarsh.Insert(fCont_Ellipse.EMPRESA, fCont_Ellipse.RESP_CONTRAT, fCont_Ellipse.D_ANUAL, fCont_Ellipse.F_INICIO_VIG _
        '                        , fCont_Ellipse.F_TERM_VIG, fCont_Ellipse.N_POLIZA, fCont_Ellipse.RUT_CONT, fCont_Ellipse.DIG_RUT_CONT _
        '                        , fCont_Ellipse.R_SOC_CONT, fCont_Ellipse.DIR_CONT, fCont_Ellipse.CIUD_CONT, fCont_Ellipse.NOMB_CONTACT _
        '                        , fCont_Ellipse.COD_CONTRAT, fCont_Ellipse.MONT_CONTRAT, fCont_Ellipse.TIP_CONTRAT, fCont_Ellipse.DESCRIP_TRAB _
        '                        , fCont_Ellipse.DISTRITO, "C", "00", Format(Today, "yyyy-MM-dd"), fCont_Ellipse.EMAIL_RESPONSABLE_CONT)
        '        cContratos_Nuevos += 1


        '    Else
        '        If (dtSMarsh.Existe_Identico(fCont_Ellipse.DISTRITO, fCont_Ellipse.COD_CONTRAT, fCont_Ellipse.D_ANUAL, fCont_Ellipse.F_INICIO_VIG, fCont_Ellipse.F_TERM_VIG _
        '                                     , fCont_Ellipse.N_POLIZA, fCont_Ellipse.RUT_CONT, fCont_Ellipse.DIG_RUT_CONT, fCont_Ellipse.MONT_CONTRAT _
        '                                     , fCont_Ellipse.DESCRIP_TRAB, fCont_Ellipse.EMPRESA, fCont_Ellipse.RESP_CONTRAT, fCont_Ellipse.DIR_CONT _
        '                                     , fCont_Ellipse.R_SOC_CONT, fCont_Ellipse.CIUD_CONT, fCont_Ellipse.NOMB_CONTACT, fCont_Ellipse.TIP_CONTRAT, "C") = 0) Then
        '            Dim Version As String
        '            Version = dtSMarsh.GetUltimaVersion(fCont_Ellipse.DISTRITO, fCont_Ellipse.COD_CONTRAT, "C")
        '            Dim iVersion As Integer
        '            iVersion = CInt(Version)
        '            iVersion = iVersion + 1
        '            Version = "0" & CStr(iVersion)
        '            Version = Strings.Right(Version, 2)
        '            dtSMarsh.Insert(fCont_Ellipse.EMPRESA, fCont_Ellipse.RESP_CONTRAT, fCont_Ellipse.D_ANUAL, fCont_Ellipse.F_INICIO_VIG _
        '                        , fCont_Ellipse.F_TERM_VIG, fCont_Ellipse.N_POLIZA, fCont_Ellipse.RUT_CONT, fCont_Ellipse.DIG_RUT_CONT _
        '                        , fCont_Ellipse.R_SOC_CONT, fCont_Ellipse.DIR_CONT, fCont_Ellipse.CIUD_CONT, fCont_Ellipse.NOMB_CONTACT _
        '                        , fCont_Ellipse.COD_CONTRAT, fCont_Ellipse.MONT_CONTRAT, fCont_Ellipse.TIP_CONTRAT, fCont_Ellipse.DESCRIP_TRAB _
        '                        , fCont_Ellipse.DISTRITO, "C", Version, Format(Today, "yyyy-MM-dd"), fCont_Ellipse.EMAIL_RESPONSABLE_CONT)
        '            cContratos_Version_Nuevos += 1
        '            'En cualquier otro caso, no existen cambios

        '        End If
        '    End If
        'Next

        ' CON CONTRATOS
        '        SendMailAdmin("Se ingresaron " & cOrdenes_Nuevas & " Ordenes nuevas <br>Se ingresaron " & cOrdenes_Version_nuevas & "  nuevas Versiones de Ordenes ya aseguradas." & _
        '                             "<br><br>Se ingresaron " & cContratos_Nuevos & " nuevos contratos <br>Se ingresaron " & cContratos_Version_Nuevos & "  nuevas Versiones de Contratos ya aseguradas.", "Interfaz MARSH:: Ingreso de Nuevas Ordenes")

        'SIN CONTRATOS
        SendMailAdmin("Se ingresaron " & cOrdenes_Nuevas & " Ordenes nuevas <br>Se ingresaron " & cOrdenes_Version_nuevas & "  nuevas Versiones de Ordenes ya aseguradas.", "Interfaz MARSH:: Ingreso de Nuevas Ordenes")
        '********************Fin  CONTRATOS+++++++++++++++++++*********************************

        'Console.WriteLine(pDistrito & "La carga necesita de las Querys de Carlos mundaca Mientras tanto se utiliza la carga directa de BD")
    End Sub


    'Procedimiento para Conectar la aplicación al MARSH y obtener los seguros para las OC o Contratos
    'Este recibe un XML y devuelve un String con Formato de XML
    Public Function CargaMARSH(ByVal pXML As String) As String
        Dim client As New WebClient()

        '**************************
        'Configuracion del Proxy
        '**************************
        Dim TablaConfiguracion As New SegurosMarshTableAdapters.SegurosMARSH_ConfiguracionTableAdapter
        Dim proxy As New WebProxy
        Dim proxyAddressAndPort, proxyUserName, proxyPassword As String
        proxyAddressAndPort = TablaConfiguracion.GetValor("CONFIGURACION", "proxyAddressAndPort")
        proxyUserName = TablaConfiguracion.GetValor("CONFIGURACION", "proxyUserName")
        proxyPassword = TablaConfiguracion.GetValor("CONFIGURACION", "proxyPassword")
        Dim cred As ICredentials
        cred = New NetworkCredential(proxyUserName, proxyPassword)
        proxy = New WebProxy(proxyAddressAndPort, True, Nothing, cred)
        WebRequest.DefaultWebProxy = proxy
        client.Proxy = proxy
        '**************************
        '**************************

        '**************************
        'Configuracion de las cabeceras y encodes
        '**************************
        client.Encoding = System.Text.Encoding.UTF8
        '  client.Headers.Add(HttpRequestHeader.ContentType, "text/html")
        client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)")

        '**************************
        '**************************
        Dim s, Address As String
        Dim s2 As Byte()
        Dim NC As New System.Collections.Specialized.NameValueCollection

        NC.Add("xml_content", pXML)
        Address = TablaConfiguracion.GetValor("CONFIGURACION", "DireccionHTTP")
        s2 = client.UploadValues(Address, NC)
        Console.WriteLine(s2)

        'Transformación de la Data de Byte a STRING
        Dim ASCIIEncoding As System.Text.ASCIIEncoding = New System.Text.ASCIIEncoding
        s = ASCIIEncoding.GetString(s2)
        Return s
    End Function


End Module
