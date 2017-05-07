Imports System.Net
Imports System.IO
Imports InterfazSegurosMARSH.Codigos


Module Module1
    Public msgDebug As String

    Public Sub Main(ByVal args() As String)

        '**************************
        'Carga de Los contratos
        'Mientras se piensa pro distrito
        '**************************
        '    CargaOC_COntratos("CMLB")
        msgDebug = PrintLogCorreo("MAIN", "Antes de Cargar OC Contratos")
        CargaOC_COntratos("ALTO")
        '**************************

        '**************************
        'Marca todos los contratos con alguna versión mayor que no se ha asegurado
        Dim TablaSeguros As New SegurosMarshTableAdapters.SegurosMARSH_SeguroTableAdapter
        msgDebug += PrintLogCorreo("MAIN", "Antes de Insertar Inoperantes")

       
        TablaSeguros.InsertarInoperantes()
        '***************************************

        '**************************
        'Procedimiento para Crear los Seguros correspondientes.
        msgDebug += PrintLogCorreo("MAIN", "Antes de Crear Seguros")
        CrearSeguros()
        '***************************************

        '**************************
        'Procedimiento que envia los avisos correspondientes
        msgDebug += PrintLogCorreo("MAIN", "Antes de enviar avisos")
        Send_Avisos_Seguros_OC()
        '***************************************
        Dim TablaConfiguracion As New SegurosMarshTableAdapters.SegurosMARSH_ConfiguracionTableAdapter
        If (TablaConfiguracion.GetValor("CONFIGURACION", "envioDebug") = "Y") Then
            SendMailAdmin(msgDebug, "DebuG MARCH", "ALTO")
        End If
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



        msgDebug += PrintLogCorreo("CrearSeguros", "OC:")
        msgDebug += PrintLogCorreo("CrearSeguros", "Antes: Se obtiene OC sin Seguros")
        '****************************
        'Creacion de ORdenes de Compra!!!!!!!!!!!
        tContratos = TablaContratos.GetOCSinSeguro
        msgDebug += PrintLogCorreo("CrearSeguros", "Cantidad de Seguros: " & tContratos.Count)
        If tContratos.Count > 0 Then
            'Luego de tener todos los contratos sin seguros, tengo que ir verificando, si estos nuevos contratos
            ' tienen una versión anterior, pero asegurada. POR CADA UNO
            'Esto será dentro del GET_XML_CONTRATO
            For Each fContrato In tContratos
                strXML = Get_XML_CONTRATO(fContrato.Cod_contrat, fContrato.Distrito, fContrato.Version, fContrato.TipoContr)

                ' Solo se ejecuta si el valor es > 0 o la fecha viene cambiada
                Dim MontoSeguroAnterior As Double
                Dim nVersion As String
                nVersion = CStr(CInt(fContrato.Version) - 1)
                nVersion = "0" & CStr(nVersion)
                nVersion = Strings.Right(nVersion, 2)
                MontoSeguroAnterior = TablaContratos.GetMaximoValorOC_Cont(fContrato.Distrito, fContrato.Cod_contrat, fContrato.TipoContr, nVersion)
                Dim fechaVencAnt As Date
                Dim FechaPars As String
                FechaPars = TablaContratos.GetUltimaFechaVencOc_Cont(fContrato.Distrito, fContrato.Cod_contrat, fContrato.TipoContr, nVersion)
                If FechaPars = "" Then FechaPars = "2011-01-01"
                fechaVencAnt = CDate(FechaPars)
                Dim diasDiferencias As Long
                diasDiferencias = DateDiff(DateInterval.Day, fechaVencAnt, CDate(fContrato.F_term_vig))
                Dim VersionSeguro As String = TablaSeguros.GetUltimaVersionSeguro(fContrato.Distrito, fContrato.Cod_contrat, fContrato.Tip_contrat)
                If (MontoSeguroAnterior < fContrato.Mont_contrat) Or (diasDiferencias > 0) Or (VersionSeguro = Nothing) Then

                    msgDebug += PrintLogCorreo("CrearSeguros", "Antes: Justo antes de Ejecutar Carga MARSH")
                    msgDebug += PrintLogCorreo("CrearSeguros", "XML a Enviar es: <br>" & Chr(10) & Chr(13) & strXML)
                    'Aqui se envia a Cargar el XML
                    ''
                 
                    strXMLResponse = CargaMARSH(strXML)
                    '
                   

                    msgDebug += PrintLogCorreo("CrearSeguros", "XML a Respuesta es:: <br>" & Chr(10) & Chr(13) & strXMLResponse)

                    If strXMLResponse = "" Then GoTo Fin

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
                    '  Dim VersionSeguro As String
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
                Else
                    'Dentro del Else va todo lo que hago para que no tome los valores que no deben tener seguros.
                    'Dim VersionSeguro As String
                    VersionSeguro = GetVersionSeguro(fContrato.Cod_contrat, fContrato.Distrito, fContrato.Version, fContrato.TipoContr)
                    VersionSeguro = Right(VersionSeguro, 2)
                    If VersionSeguro = "" Then
                        VersionSeguro = "00"
                    End If
                    TablaSeguros.Insert(fContrato.Cod_contrat, _
                                        fContrato.Distrito, fContrato.TipoContr, fContrato.Version, _
                                        "NN" & fContrato.Cod_contrat, _
                                        "", _
                                        "", _
                                        0, _
                                        0, _
                                        "NO", _
                                        "Monto Menor al anterior o Cero", _
                                        "", "sinAviso@xstratacopper.cl", "", Format(Today, "yyyy-MM-dd"), VersionSeguro)
                End If
Fin:
            Next
        End If


        '****************************
        'Creacion de Contratos!!!!!!!!!!!
        msgDebug += PrintLogCorreo("CrearSeguros", "Creacion los Contratos")

        tContratos = TablaContratos.GetContratosSinSeguro
        msgDebug += PrintLogCorreo("CrearSeguros", "Cantidad de Seguros para Contratos: " & tContratos.Count)

        If tContratos.Count > 0 Then

            For Each fContrato In tContratos
                strXML = Get_XML_CONTRATO(fContrato.Cod_contrat, fContrato.Distrito, fContrato.Version, fContrato.TipoContr)

                ' Solo se ejecuta si el valor es > 0 o la fecha viene cambiada
                Dim MontoSeguroAnterior As Double
                Dim nVersion As String
                nVersion = CStr(CInt(fContrato.Version) - 1)
                nVersion = "0" & CStr(nVersion)
                nVersion = Strings.Right(nVersion, 2)
                MontoSeguroAnterior = TablaContratos.GetMaximoValorOC_Cont(fContrato.Distrito, fContrato.Cod_contrat, fContrato.TipoContr, nVersion)
                Dim fechaVencAnt As Date
                Dim FechaPars As String
                FechaPars = TablaContratos.GetUltimaFechaVencOc_Cont(fContrato.Distrito, fContrato.Cod_contrat, fContrato.TipoContr, nVersion)
                If FechaPars = "" Then FechaPars = "2011-01-01"
                fechaVencAnt = CDate(FechaPars)
                Dim diasDiferencias As Long
                diasDiferencias = DateDiff(DateInterval.Day, fechaVencAnt, CDate(fContrato.F_term_vig))
                Dim VersionSeguro As String = TablaSeguros.GetUltimaVersionSeguro(fContrato.Distrito, fContrato.Cod_contrat, fContrato.Tip_contrat)
                If (MontoSeguroAnterior < fContrato.Mont_contrat) Or (diasDiferencias > 0) Or (VersionSeguro = Nothing) Then

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
                Else
                    'Dentro del Else va todo lo que hago para que no tome los valores que no deben tener seguros.
                    VersionSeguro = GetVersionSeguro(fContrato.Cod_contrat, fContrato.Distrito, fContrato.Version, fContrato.TipoContr)
                    If VersionSeguro = "" Then
                        VersionSeguro = "00"
                    End If
                    VersionSeguro = Right(VersionSeguro, 2)
                    TablaSeguros.Insert(fContrato.Cod_contrat, _
                                        fContrato.Distrito, fContrato.TipoContr, fContrato.Version, _
                                        "NN" & fContrato.Cod_contrat, _
                                        "", _
                                        "", _
                                        0, _
                                        0, _
                                        "NO", _
                                        "Monto Menor al anterior o Cero", _
                                        "", "sinAviso@xstratacopper.cl", "", Format(Today, "yyyy-MM-dd"), VersionSeguro)
                End If
            Next
        End If
        msgDebug += PrintLogCorreo("CrearSeguros", "FIN")

    End Sub
    'Procedimiento para Avisar ingresos
    

    Public Sub Send_Avisos_Seguros_OC()
        Dim Mensage, MensageOC, MensageContrato, MensageAux, EmailSeguro, Distrito As String
        Dim taSegurosSinAviso As New SegurosMarshTableAdapters.SegurosMARSH_SeguroTableAdapter
        Dim tSeguros As SegurosMarsh.SegurosMARSH_SeguroDataTable
        Dim taContratosOC As New SegurosMarshTableAdapters.SegurosMARSH_Oc_ContratoTableAdapter
        tSeguros = taSegurosSinAviso.GetSegurosPorCorreo
        ' tSeguros = taSegurosSinAviso.GetDataSeguro
        Dim tseguro As SegurosMarsh.SegurosMARSH_SeguroRow
        Dim EmailAlto, EmailCMLB As New ArrayList
        Dim iContratos, iOrdenes As Integer
        Distrito = ""

        'Dim eeemail As Array
        'Email.Resize(Email, 3)
        Dim TablaConfiguracion As New SegurosMarshTableAdapters.SegurosMARSH_ConfiguracionTableAdapter()
        'Email.AddRange(TablaConfiguracion.GetCorreos("Admin").GetEnumerator)
        Dim x As Integer = 1
        For Each Correo As SegurosMarsh.SegurosMARSH_ConfiguracionRow In TablaConfiguracion.GetCorreos("AdminALTO")
            EmailAlto.Add(Correo.Valor2)
            x += 1
        Next

        'x = 1
        'For Each Correo As SegurosMarsh.SegurosMARSH_ConfiguracionRow In TablaConfiguracion.GetCorreos("AdminCMLB")
        '    EmailCMLB.Add(Correo.Valor2)
        '    x += 1
        'Next
        EmailAlto.Insert(0, "")
        'EmailCMLB.Insert(0, "")
        EmailSeguro = ""

        Mensage = " Estimado, en la ultima actualización de seguros MARSH se cuentan los siguientes cambios: <br><br>"
        If tSeguros.Count = 0 Then
            MensageOC = "No se han encontrado seguros asociados a órdenes o contratos<br><br>"

        Else
            'Ordenes de compra
            MensageOC = "Se han ingresado las siguientes órdenes de compra: <br><br>"
            MensageOC += "<table width='100%' border = '1' > "
            MensageOC += "<tr><td><b>OC</b></td>"
            MensageOC += "<td><b>Descripcion</b></td>"
            MensageOC += "<td><b>Versión Seguro</b></td> "
            MensageOC += "<td><b>Id Verificación</b></td> "
            MensageOC += "<td><b>Fecha Inicio Declaración</b></td> "
            MensageOC += "<td><b>Fecha Fin Declaración</b></td> "
            MensageOC += "<td><b>Aprobado</b></td>"
            MensageOC += "<td><b>Motivo</b></td>"
            MensageOC += "<td><b>Monto Declarado(UF)</b></td>"
            MensageOC += "<td><b>Prima(UF)</b></td>"
            MensageOC += "<td><b>URL Certificado</b></td>"
            MensageOC += "</tr>"

            'Contratos
            MensageContrato = "Se han ingresado los siguientes contratos: <br><br>"
            MensageContrato += "<table width='100%' border = '1' > "
            MensageContrato += "<tr><td><b>OC</b></td>"
            MensageContrato += "<td><b>Descripcion</b></td>"
            MensageContrato += "<td><b>Versión Seguro</b></td> "
            MensageContrato += "<td><b>Id Verificación</b></td> "
            MensageContrato += "<td><b>Fecha Inicio Declaración</b></td> "
            MensageContrato += "<td><b>Fecha Fin Declaración</b></td> "
            MensageContrato += "<td><b>Aprobado</b></td>"
            MensageContrato += "<td><b>Motivo</b></td>"
            MensageContrato += "<td><b>Monto Declarado(UF)</b></td>"
            MensageContrato += "<td><b>Prima(UF)</b></td>"
            MensageContrato += "<td><b>URL Certificado</b></td>"
            MensageContrato += "</tr>"




            For Each tseguro In tSeguros


                If EmailSeguro = "" Then
                    EmailSeguro = tseguro.EmailAviso
                End If

                If EmailSeguro <> tseguro.EmailAviso Then

                    If iContratos > 0 Then
                        Mensage += MensageContrato
                        Mensage += "</Table>"
                    End If

                    If iOrdenes > 0 Then
                        Mensage += MensageOC
                        Mensage += "</Table>"
                    End If

                    If (Distrito = "CMLB") Then
                        'EmailCMLB(0) = EmailSeguro
                        'SendMail(EmailCMLB, Mensage, "Interfaz MARSH:: Creación de Nuevos Seguros")
                        'EmailAlto(0) = ""
                    Else
                        EmailAlto(0) = EmailSeguro
                        SendMail(EmailAlto, Mensage, "Interfaz MARSH:: Creación de Nuevos Seguros")
                        'EmailCMLB(0) = ""
                    End If
                    'taSegurosSinAviso.ActualizarEnviados()
                    Distrito = tseguro.Distrito

                    EmailSeguro = tseguro.EmailAviso
                    Mensage = " Estimado, en la ultima actualización de seguros MARSH se cuentan los siguientes cambios: <br><br>"

                    iContratos = 0
                    iOrdenes = 0

                    MensageOC = "<br>Se han ingresado las siguientes órdenes de compra: <br><br>"
                    MensageOC += "<table width='100%' border = '1' > "
                    MensageOC += "<tr><td><b>OC</b></td>"
                    MensageOC += "<td><b>Descripcion</b></td>"
                    MensageOC += "<td><b>Versión Seguro</b></td> "
                    MensageOC += "<td><b>Id Verificación</b></td> "
                    MensageOC += "<td><b>Fecha Inicio Declaración</b></td> "
                    MensageOC += "<td><b>Fecha Fin Declaración</b></td> "
                    MensageOC += "<td><b>Aprobado</b></td>"
                    MensageOC += "<td><b>Motivo</b></td>"
                    MensageOC += "<td><b>Monto Declarado(UF)</b></td>"
                    MensageOC += "<td><b>Prima(UF)</b></td>"
                    MensageOC += "<td><b>URL Certificado</b></td>"
                    MensageOC += "</tr>"

                    MensageContrato = "Se han ingresado los sigientes contratos: <br><br>"
                    MensageContrato += "<table width='100%' border = '1' > "
                    MensageContrato += "<tr><td><b>OC</b></td>"
                    MensageContrato += "<td><b>Descripcion</b></td>"
                    MensageContrato += "<td><b>Versión Seguro</b></td> "
                    MensageContrato += "<td><b>Id Verificación</b></td> "
                    MensageContrato += "<td><b>Fecha Inicio Declaración</b></td> "
                    MensageContrato += "<td><b>Fecha Fin Declaración</b></td> "
                    MensageContrato += "<td><b>Aprobado</b></td>"
                    MensageContrato += "<td><b>Motivo</b></td>"
                    MensageContrato += "<td><b>Monto Declarado(UF)</b></td>"
                    MensageContrato += "<td><b>Prima(UF)</b></td>"
                    MensageContrato += "<td><b>URL Certificado</b></td>"
                    MensageContrato += "</tr>"

                End If
                MensageAux = "<tr><td><b>" & tseguro.Cod_contrat & " </b></td>"
                MensageAux += "<td><b>" & taContratosOC.GetDescripcionOC_COnt(tseguro.Cod_contrat, tseguro.Distrito, tseguro.TipoContr, tseguro.Version) & " </b></td>"
                MensageAux += "<td><b>" & tseguro.Version_Seguro & "</b></td> "
                MensageAux += "<td><b>" & tseguro.Id_verificacion & "</b></td> "
                MensageAux += "<td><b>" & tseguro.F_inicio_declaracion & "</b></td> "
                MensageAux += "<td><b>" & tseguro.F_term_declaracion & "</b></td> "
                MensageAux += "<td><b>" & tseguro.Aprobado & "</b></td>"
                MensageAux += "<td><b>" & tseguro.Motivo & "</b></td>"
                MensageAux += "<td><b>" & tseguro.Mont_declarado & "</b></td>"
                MensageAux += "<td><b>" & tseguro.Prima_contrat & "</b></td>"
                MensageAux += "<td><b>" & tseguro.URL_Certificado & "</b></td>"
                MensageAux += "</tr>"

                If tseguro.TipoContr = "C" Then
                    MensageContrato += MensageAux
                    iContratos += 1
                Else
                    MensageOC += MensageAux
                    iOrdenes += 1
                End If
                Distrito = tseguro.Distrito
            Next
            If iContratos > 0 Then
                Mensage += MensageContrato
                Mensage += "</Table>"
            End If

            If iOrdenes > 0 Then
                Mensage += MensageOC
                Mensage += "</Table>"
            End If
          
            If (Distrito = "CMLB") Then
                'EmailCMLB(0) = EmailSeguro
                'SendMail(EmailCMLB, Mensage, "Interfaz MARSH:: Creación de Nuevos Seguros")
                ' EmailAlto(0) = ""
            Else
                EmailAlto(0) = EmailSeguro
                SendMail(EmailAlto, Mensage, "Interfaz MARSH:: Creación de Nuevos Seguros")
                'EmailCMLB(0) = ""
            End If

            taSegurosSinAviso.ActualizarEnviados()


            ''Contratos



        End If

    End Sub


    'Procedimiento que Entrega el XML de una fila identificada.
    Public Function Get_XML_CONTRATO(ByVal Contra_OC As String, ByVal pDistrito As String, ByVal pVersion As String, ByVal pTipo As String) As String
        Dim TablaContratos As SegurosMarshTableAdapters.SegurosMARSH_Oc_ContratoTableAdapter
       
        TablaContratos = New SegurosMarshTableAdapters.SegurosMARSH_Oc_ContratoTableAdapter

        Dim VersionSeguro As String
        'Dim FechaVencSeguroAnterior As Date
        Dim MontoSeguroAnterior, MontoSeguroActual As Double
        msgDebug += PrintLogCorreo("Get_XML_CONTRATO", "Obteniendo Version")
        VersionSeguro = GetVersionSeguro(Contra_OC, pDistrito, pVersion, pTipo)
        MontoSeguroActual = TablaContratos.GetUltimoValorOC_Cont(pDistrito, Contra_OC, pTipo, pVersion)
        MontoSeguroAnterior = 0
        If VersionSeguro <> "" Then
            Dim nVersion As String
            nVersion = CStr(CInt(pVersion) - 1)
            nVersion = "0" & CStr(nVersion)
            nVersion = Strings.Right(nVersion, 2)
            MontoSeguroAnterior = TablaContratos.GetMaximoValorOC_Cont(pDistrito, Contra_OC, pTipo, nVersion)
            If MontoSeguroAnterior = Nothing Then
                MontoSeguroAnterior = 0
            End If

        End If

        'Monto Seguro no tiene que ser negativo
        
        If MontoSeguroAnterior = 0 Then
            Return (TablaContratos.Get_XML_CONTRATO_OC(VersionSeguro, 0, Contra_OC, pDistrito, pTipo, pVersion))
        Else
            If MontoSeguroActual < MontoSeguroAnterior Then
                MontoSeguroAnterior = MontoSeguroActual
            End If
            Return (TablaContratos.Get_XML_CONTRATO_OC(VersionSeguro, MontoSeguroAnterior, Contra_OC, pDistrito, pTipo, pVersion))
        End If


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
        Dim sOrdenes_Nuevas, sOrdenes_Version_nuevas, sContratos_Nuevos, sContratos_Version_Nuevos As String
        Dim Fecha_Termino, Fecha_inicio As String
        sOrdenes_Nuevas = ""
        sOrdenes_Version_nuevas = ""
        sContratos_Nuevos = ""
        sContratos_Version_Nuevos = ""
        '
        Dim dtOC_Ellipse As New SegurosMarshTableAdapters.Ellipse_OCTableAdapter
        Dim fOC_Ellipse As SegurosMarsh.Ellipse_OCRow
        Dim dOC_Ellipse As SegurosMarsh.Ellipse_OCDataTable

        Dim TablaConfiguracion As New SegurosMarshTableAdapters.SegurosMARSH_ConfiguracionTableAdapter

        '************************ Configuracion de Límite de fecha seguro*********************************
        Dim Año As Integer
        Año = Year(Today)
        Dim FechaVencimientoPoliza As Date

        'Obtencion de Fecha Vencimiento Poliza (año,mes,dia) (año,valor,valor2) del año en curso
        FechaVencimientoPoliza = DateAndTime.DateSerial(Año, TablaConfiguracion.GetValor("CONSTANTES", "FECHA_FIN_POLISA"), TablaConfiguracion.GetValor2("CONSTANTES", "FECHA_FIN_POLISA"))

        'en caso que la fecha de vencimiento sea mayor a hoy, se cambia año por el año anteior
        If Today <= FechaVencimientoPoliza Then

            Año = Año - 1
        End If
        FechaVencimientoPoliza = DateAndTime.DateSerial(Año, TablaConfiguracion.GetValor("CONSTANTES", "FECHA_FIN_POLISA"), TablaConfiguracion.GetValor2("CONSTANTES", "FECHA_FIN_POLISA"))
        '*****************************************************************************
        msgDebug += PrintLogCorreo("CargaOC_COntratos", "Fecha Vencimiento Poliza: " & FechaVencimientoPoliza)

        Dim dtSMarsh As New SegurosMarshTableAdapters.SegurosMARSH_Oc_ContratoTableAdapter

        '******************** ORDENES DE COMPRAAAAAAAAA*********************************
        ' se ejecuta la siguiente Query:
        'Select Case EMPRESA, RESP_CONTRAT, D_ANUAL, F_INICIO_VIG, F_TERM_VIG, N_POLIZA, RUT_CONT, DIG_RUT_CONT, R_SOC_CONT, DISTRITO, DIR_CONT, CIUD_CONT, 
        'NOMB_CONTACT, COD_CONTRAT, MONT_CONTRAT, TIP_CONTRAT, DESCRIP_TRAB, MONT_ORDEN_UF, RESPONSABLE_CONT,
        'EMAIL_RESPONSABLE_CONT
        'From XSV220_SEGUROS_MARSH
        'Where(DISTRITO = : DISTRITO)
        ' VIsta se comporta de la siguiente manera::::XSV220_SEGUROS_MARSH

        'Ojo que siemrpe trae los mismos dependiendo de cuantos días hacia atras este seteada la Vista en este caso son 6 dias.
        'TO_CHAR(SYSDATE - 6, 'yyyyMMdd'))
        '  Or (A.ORDER_DATE >= TO_CHAR (SYSDATE - 6, 'yyyyMMdd')
        msgDebug += PrintLogCorreo("CargaOC_COntratos", "Antes de obtener las Ordenes de Ellipse")

        dOC_Ellipse = dtOC_Ellipse.GetOrdenes_compra_Distrito(pDistrito)

        msgDebug += PrintLogCorreo("CargaOC_COntratos", "Cantidad de Ordenes obtenidas desde Ellipse:" & dOC_Ellipse.Count)

        For Each fOC_Ellipse In dOC_Ellipse

            Dim MontoContrato As Double
            'Monto Minimo de Contrato es de 2 UF 
            '.- 20110920 .- Modificado Monto Contrato = 20
            MontoContrato = CDbl(TablaConfiguracion.GetValor("CONSTANTES", "MINIMO_UF_OC"))
            If CDbl(fOC_Ellipse.MONT_CONTRAT) <= MontoContrato Then
                If CDbl(fOC_Ellipse.MONT_ORDEN_UF) >= MontoContrato Then
                    MontoContrato = fOC_Ellipse.MONT_ORDEN_UF
                End If
            Else
                If fOC_Ellipse.MONT_ORDEN_UF >= fOC_Ellipse.MONT_CONTRAT Then
                    MontoContrato = fOC_Ellipse.MONT_ORDEN_UF
                Else
                    MontoContrato = fOC_Ellipse.MONT_CONTRAT
                End If

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
            Dim Dias As Integer
            Dias = CInt(TablaConfiguracion.GetValor("CONSTANTES", "FECHA_TERMINO_ORDEN"))
            Fecha_Termino = Format(DateAdd(DateInterval.Day, Dias, CDate(fOC_Ellipse.F_TERM_VIG)), "yyyy-MM-dd")
            If (DateDiff(DateInterval.Day, CDate(fOC_Ellipse.F_INICIO_VIG), Now) < 0) Then
                Fecha_inicio = Format(Now, "yyyy-MM-dd")
            Else
                If (CDate(fOC_Ellipse.F_INICIO_VIG) <= FechaVencimientoPoliza) Then
                    Fecha_inicio = Format(FechaVencimientoPoliza, "yyyy-MM-dd")
                Else
                    Fecha_inicio = Format(CDate(fOC_Ellipse.F_INICIO_VIG), "yyyy-MM-dd")
                End If


                End If
                Dim Ciud_Contratista As String = ""
                If Not fOC_Ellipse.IsCIUD_CONTNull Then
                    Ciud_Contratista = fOC_Ellipse.CIUD_CONT
                End If
                Dim nNOMB_CONTACT As String = "N/A"
                If Not fOC_Ellipse.IsNOMB_CONTACTNull Then
                    nNOMB_CONTACT = fOC_Ellipse.NOMB_CONTACT
                End If
                Dim nR_SOC_CONT As String = ""
                If Not fOC_Ellipse.IsR_SOC_CONTNull Then
                    nR_SOC_CONT = fOC_Ellipse.R_SOC_CONT
                End If
                Dim nDIR_CONT As String = ""
                If Not fOC_Ellipse.IsDIR_CONTNull Then
                    nDIR_CONT = fOC_Ellipse.DIR_CONT
                End If
            'Si no existe Orden de compra, se inserta en tabla, con version 00
            If (dtSMarsh.Existe_OC_Contrato(fOC_Ellipse.DISTRITO, "O", fOC_Ellipse.COD_CONTRAT) = 0) Then

                dtSMarsh.Insert(fOC_Ellipse.EMPRESA, fOC_Ellipse.RESP_CONTRAT, fOC_Ellipse.D_ANUAL, Fecha_inicio _
                                , Fecha_Termino, fOC_Ellipse.N_POLIZA, fOC_Ellipse.RUT_CONT, fOC_Ellipse.DIG_RUT_CONT _
                                , nR_SOC_CONT, nDIR_CONT, Ciud_Contratista, nNOMB_CONTACT _
                                , fOC_Ellipse.COD_CONTRAT, MontoContrato, fOC_Ellipse.TIP_CONTRAT, fOC_Ellipse.DESCRIP_TRAB _
                                , fOC_Ellipse.DISTRITO, "O", "00", Format(Today, "yyyy-MM-dd"), fOC_Ellipse.EMAIL_RESPONSABLE_CONT)
                cOrdenes_Nuevas += 1
                sOrdenes_Nuevas += fOC_Ellipse.COD_CONTRAT & ";"


            Else
                'Si no existe identico segun parámetros más abajo descritos se agrega una version mñas despues de haber encontrado la ultima version
                'Segun Requerimieo de Carlos Mundaca es necesario que no se Cree.
                '- Cambio realizado 2016-07-28
                'If (dtSMarsh.Existe_Identico_Asegurado(fOC_Ellipse.DISTRITO, fOC_Ellipse.COD_CONTRAT, Fecha_Termino, MontoContrato, "O") = 0) Then
                If (dtSMarsh.Existe_Identico(fOC_Ellipse.DISTRITO, fOC_Ellipse.COD_CONTRAT, Fecha_Termino, MontoContrato, fOC_Ellipse.TIP_CONTRAT, "O", fOC_Ellipse.RESP_CONTRAT, fOC_Ellipse.N_POLIZA) = 0) Then
                    Dim Version As String
                    Version = dtSMarsh.GetUltimaVersion(fOC_Ellipse.DISTRITO, fOC_Ellipse.COD_CONTRAT, "O")

                    'Nuevo Cambio ***********************


                    Dim iVersion As Integer
                    iVersion = CInt(Version)
                    iVersion = iVersion + 1
                    Version = "0" & CStr(iVersion)
                    Version = Strings.Right(Version, 2)
                    dtSMarsh.Insert(fOC_Ellipse.EMPRESA, fOC_Ellipse.RESP_CONTRAT, fOC_Ellipse.D_ANUAL, Fecha_inicio _
                                , Fecha_Termino, fOC_Ellipse.N_POLIZA, fOC_Ellipse.RUT_CONT, fOC_Ellipse.DIG_RUT_CONT _
                                , nR_SOC_CONT, nDIR_CONT, Ciud_Contratista, nNOMB_CONTACT _
                                , fOC_Ellipse.COD_CONTRAT, MontoContrato, fOC_Ellipse.TIP_CONTRAT, fOC_Ellipse.DESCRIP_TRAB _
                                , fOC_Ellipse.DISTRITO, "O", Version, Format(Today, "yyyy-MM-dd"), fOC_Ellipse.EMAIL_RESPONSABLE_CONT)
                    cOrdenes_Version_nuevas += 1
                    sOrdenes_Version_nuevas += fOC_Ellipse.COD_CONTRAT & ";"
                    'En cualquier otro caso, no existen cambios
                End If
            End If
        Next
        '******************** FIN   ORDENES DE COMPRAAAAAAAAA*********************************

        msgDebug += PrintLogCorreo("CargaOC_COntratos", "Termino de Agregar OC, comenzando con Contratos")

        '******************** CONTRATOS+++++++++++++++++++*********************************

        Dim dtCont_Ellipse As New SegurosMarshTableAdapters.Ellipse_ContratosTableAdapter
        Dim fCont_Ellipse As SegurosMarsh.Ellipse_ContratosRow
        Dim dCont_Ellipse As SegurosMarsh.Ellipse_ContratosDataTable


        dCont_Ellipse = dtCont_Ellipse.GetContratos_Distrito(pDistrito)
        msgDebug += PrintLogCorreo("CargaOC_COntratos", "Cantidad de Contratos obtenidos desde Ellipse:" & dOC_Ellipse.Count)
        For Each fCont_Ellipse In dCont_Ellipse
            Dim Dias As Integer
            Dias = CInt(TablaConfiguracion.GetValor("CONSTANTES", "FECHA_TERMINO_CONTRATO"))
            Fecha_Termino = Format(DateAdd(DateInterval.Day, Dias, CDate(fCont_Ellipse.F_TERM_VIG)), "yyyy-MM-dd")

            Dim DAnual As String
            If (DateDiff(DateInterval.Day, CDate(fCont_Ellipse.F_INICIO_VIG), Now) < 0) Then
                Fecha_inicio = Format(Now, "yyyy-MM-dd")
            Else
                If (CDate(fCont_Ellipse.F_INICIO_VIG) <= FechaVencimientoPoliza) Then
                    Fecha_inicio = Format(FechaVencimientoPoliza, "yyyy-MM-dd")
                Else
                    Fecha_inicio = Format(CDate(fCont_Ellipse.F_INICIO_VIG), "yyyy-MM-dd")
                End If

            End If
            If DateDiff(DateInterval.Month, CDate(fCont_Ellipse.F_INICIO_VIG), CDate(Fecha_Termino)) >= 12 Then
                DAnual = "Si"
            Else
                DAnual = "No"

            End If
            Dim Ciud_Contratista As String = ""
            If Not fCont_Ellipse.IsCIUD_CONTNull Then
                Ciud_Contratista = fCont_Ellipse.CIUD_CONT
            End If
            Dim NombreContacto As String = ""
            If Not fCont_Ellipse.IsNOMB_CONTACTNull Then
                NombreContacto = fCont_Ellipse.NOMB_CONTACT
            End If
            Dim nNOMB_CONTACT As String = "N/A"
            If Not fCont_Ellipse.IsNOMB_CONTACTNull Then
                nNOMB_CONTACT = nNOMB_CONTACT
            End If
            Dim nR_SOC_CONT As String = ""
            If Not fCont_Ellipse.IsR_SOC_CONTNull Then
                nR_SOC_CONT = fCont_Ellipse.R_SOC_CONT
            End If
            Dim nDIR_CONT As String = ""
            If Not fCont_Ellipse.IsDIR_CONTNull Then
                nDIR_CONT = fCont_Ellipse.DIR_CONT
            End If

            If (dtSMarsh.Existe_OC_Contrato(fCont_Ellipse.DISTRITO, "C", fCont_Ellipse.COD_CONTRAT) = 0) Then
                dtSMarsh.Insert(fCont_Ellipse.EMPRESA, fCont_Ellipse.RESP_CONTRAT, DAnual, Fecha_inicio _
                                , Fecha_Termino, fCont_Ellipse.N_POLIZA, fCont_Ellipse.RUT_CONT, fCont_Ellipse.DIG_RUT_CONT _
                                , nR_SOC_CONT, nDIR_CONT, Ciud_Contratista, NombreContacto _
                                , fCont_Ellipse.COD_CONTRAT, fCont_Ellipse.MONT_CONTRAT, fCont_Ellipse.TIP_CONTRAT, fCont_Ellipse.DESCRIP_TRAB _
                                , fCont_Ellipse.DISTRITO, "C", "00", Format(Today, "yyyy-MM-dd"), fCont_Ellipse.EMAIL_RESPONSABLE_CONT)
                cContratos_Nuevos += 1
                sContratos_Nuevos += fCont_Ellipse.COD_CONTRAT & ";"

            Else
                'Si no existe identico segun parámetros más abajo descritos se agrega una version mñas despues de haber encontrado la ultima version
                'Segun Requerimieo de Carlos Mundaca es necesario que no se Cree.
                '- Cambio realizado 2016-07-28
                'If (dtSMarsh.Existe_Identico_Asegurado(fCont_Ellipse.DISTRITO, fCont_Ellipse.COD_CONTRAT, Fecha_Termino , fCont_Ellipse.MONT_CONTRAT, "C") = 0) Then
                If (dtSMarsh.Existe_Identico(fCont_Ellipse.DISTRITO, fCont_Ellipse.COD_CONTRAT, Fecha_Termino, fCont_Ellipse.MONT_CONTRAT, fCont_Ellipse.TIP_CONTRAT, "C", fOC_Ellipse.RESP_CONTRAT, fOC_Ellipse.N_POLIZA) = 0) Then
                    Dim Version As String
                    Version = dtSMarsh.GetUltimaVersion(fCont_Ellipse.DISTRITO, fCont_Ellipse.COD_CONTRAT, "C")
                    Dim iVersion As Integer
                    iVersion = CInt(Version)
                    iVersion = iVersion + 1
                    Version = "0" & CStr(iVersion)
                    Version = Strings.Right(Version, 2)
                    dtSMarsh.Insert(fCont_Ellipse.EMPRESA, fCont_Ellipse.RESP_CONTRAT, DAnual, Fecha_inicio _
                                , Fecha_Termino, fCont_Ellipse.N_POLIZA, fCont_Ellipse.RUT_CONT, fCont_Ellipse.DIG_RUT_CONT _
                                , nR_SOC_CONT, nDIR_CONT, Ciud_Contratista, NombreContacto _
                                , fCont_Ellipse.COD_CONTRAT, fCont_Ellipse.MONT_CONTRAT, fCont_Ellipse.TIP_CONTRAT, fCont_Ellipse.DESCRIP_TRAB _
                                , fCont_Ellipse.DISTRITO, "C", Version, Format(Today, "yyyy-MM-dd"), fCont_Ellipse.EMAIL_RESPONSABLE_CONT)
                    cContratos_Version_Nuevos += 1
                    sContratos_Version_Nuevos += fCont_Ellipse.COD_CONTRAT & ";"
                    'En cualquier otro caso, no existen cambios

                End If
            End If
        Next
        msgDebug += PrintLogCorreo("CargaOC_COntratos", "Terminada la Carga de OC y Contratos, a continuacion se envia el Mail")

        ' CON CONTRATOS
        SendMailAdmin("Se ingresaron " & cOrdenes_Nuevas & " Ordenes nuevas (" & sOrdenes_Nuevas & ") <br>Se ingresaron " & cOrdenes_Version_nuevas & "  nuevas Versiones de Ordenes ya aseguradas (" & sOrdenes_Version_nuevas & ")." & _
                             "<br><br>Se ingresaron " & cContratos_Nuevos & " nuevos contratos (" & sContratos_Nuevos & ") <br>Se ingresaron " & cContratos_Version_Nuevos & "  nuevas Versiones de Contratos ya aseguradas (" & sContratos_Version_Nuevos & ").", "Interfaz MARSH PRD:: Ingreso de Nuevos Seguros", pDistrito)

            'SIN CONTRATOS
        'SendMailAdmin("Se ingresaron " & cOrdenes_Nuevas & " Ordenes nuevas <br>Se ingresaron " & cOrdenes_Version_nuevas & "  nuevas Versiones de Ordenes ya aseguradas.", "Interfaz MARSH:: Ingreso de Nuevas Ordenes")
            '********************Fin  CONTRATOS+++++++++++++++++++*********************************

            'Console.WriteLine(pDistrito & "La carga necesita de las Querys de Carlos mundaca Mientras tanto se utiliza la carga directa de BD")
    End Sub


    'Procedimiento para Conectar la aplicación al MARSH y obtener los seguros para las OC o Contratos
    'Este recibe un XML y devuelve un String con Formato de XML
    Public Function CargaMARSH(ByVal pXML As String) As String
        Dim client As New WebClient()
        Dim proxyAddressAndPort, proxyUserName, proxyPassword, Address, proxyDomain As String
        proxyAddressAndPort = ""
        proxyUserName = ""
        proxyPassword = ""
        Address = ""
        proxyDomain = ""
        Try
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3
            msgDebug += PrintLogCorreo("CargaMARSH", "Dentro del envio de Carga MARSH")

            '**************************
            'Configuracion del Proxy
            '**************************
            Dim TablaConfiguracion As New SegurosMarshTableAdapters.SegurosMARSH_ConfiguracionTableAdapter
            Dim proxy As New WebProxy

            proxyAddressAndPort = TablaConfiguracion.GetValor("CONFIGURACION", "proxyAddressAndPort")
            proxyUserName = TablaConfiguracion.GetValor("CONFIGURACION", "proxyUserName")
            proxyPassword = TablaConfiguracion.GetValor("CONFIGURACION", "proxyPassword")
            proxyDomain = TablaConfiguracion.GetValor("CONFIGURACION", "proxyDomain")
            Dim cred As ICredentials
            cred = New NetworkCredential(proxyUserName, proxyPassword, proxyDomain)
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
            Dim s As String
            Dim s2 As Byte()
            Dim NC As New System.Collections.Specialized.NameValueCollection

            NC.Add("xml_content", pXML)
            Address = TablaConfiguracion.GetValor("CONFIGURACION", "DireccionHTTP")
            'Threading.Thread.Sleep(5000)


            s2 = client.UploadValues(Address, NC)
            Console.WriteLine(s2)

            'Transformación de la Data de Byte a STRING
            Dim ASCIIEncoding As System.Text.ASCIIEncoding = New System.Text.ASCIIEncoding
            s = ASCIIEncoding.GetString(s2)
            Return s
        Catch ex As Exception
            Dim Mensage As String
            Mensage = "A ocurrido un error al cargar el seguro a la página Web <br>"
            Mensage += "Datos Utilizados son: <br>"
            Mensage += "Proxy       :: " & proxyAddressAndPort & "<br>"
            Mensage += "Usuario     :: " & proxyUserName & "<br>"
            Mensage += "Password    :: " & proxyPassword & " <br>"
            Mensage += "Dominio     :: " & proxyDomain & " <br>"
            Mensage += "URL Seguro  :: " & Address & " <br>"
            Mensage += "El error Es:: <br>"
            Mensage += ex.ToString
            Mensage += "<br><br> el XML enviado es <br>"
            Mensage += pXML
            Send_Correo(Mensage, "Interfas MARSH:: ERROR EN CARGA MARSH AL SEGURO")
            Return ""
        End Try
    End Function


    Public Sub Send_Correo(ByVal Mensage As String, ByVal Subject As String)

        Dim Email As New ArrayList


        Email.Add("alex.castillo@otispa.cl")
        Email.Add("carlos.mundaca@glencore.cl")
        'Email.Insert(0, "")



        SendMail(Email, Mensage, Subject)

    End Sub

End Module
