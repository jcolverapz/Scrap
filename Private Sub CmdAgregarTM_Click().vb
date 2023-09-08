Private Sub CmdAgregarTM_Click()
    bandera_SupScrap = False
If IdOperacion = 0 Then
    MsgBox "Seleccione una Operacion.", vbInformation, MSG
    Exit Sub
End If

    ValText "Tipo Scrap", Me.CboTipoSrap.Text
    If Band = False Then Exit Sub
    ValText "Clave", Me.TxtClaveP.Text
    If Band = False Then Exit Sub
    ValText "Concepto", Me.CboConceptoScrap.Text
    If Band = False Then Exit Sub
    ValNum "Cantidad", Me.TxtCantidad.Text
    If Band = False Then Exit Sub

CNN.CmdBuscaScrapCapturadoxOperacion (IdOperacion), (IdJC), (Turno), (Fecha)
If CNN.rsCmdBuscaScrapCapturadoxOperacion.EOF <> True Then
   TScrapxOperacion = CNN.rsCmdBuscaScrapCapturadoxOperacion!TotalScrap
Else
    TScrapxOperacion = 0
End If
CNN.rsCmdBuscaScrapCapturadoxOperacion.Close

If Me.CboTipoSrap.Text = "PROVEEDOR" Then
    TicketProv = ""
    
    TipoScrap_Scrap = "Proveedor"
Else
         If banderaTemplado = 1 Then
            TipoScrap_Scrap = "Templado"
         Else
              TipoScrap_Scrap = "Proceso"
              If Bandera_Retrabajo = "Retrabajo" Then
                    TipoScrap_Scrap = "Retrabajo"
              End If
              
         End If
      
    End If

    Dim dia As Date
    dia = (Date - 1)
    
'    If FuncionId = 15 Or FuncionId = 118 Or DptoId = "09" Or DptoId = "03" Or EmpleadoId = "001-151" Or EmpleadoId = "001-045" Or FuncionId = 105 Then
'
'     bandera_SupScrap = True
'
'    Else
'        If (Fecha) < (dia) Then
'
'            MsgBox "No se puede Registrar Scrap Fechas pasados", vbInformation, MSG
'            bandera_SupScrap = False
'
'            Exit Sub
'        End If
 '''''''''''''''''''''''''''''''''''''''''''''''''''Revisar si han pasado 30 min despues del turno para poder continuar'''''''''''''''''''''''
'        CNN.cmdTurnoScrap
'        If CNN.rscmdTurnoScrap.EOF <> True Then
'            If CNN.rscmdTurnoScrap!SalidaHoy = 1 Then
'                TiempoEspera = CNN.rscmdTurnoScrap!Tiempo_Espera
'                Call TurnoT
'                If FechaTurno <> Fecha Then
'                    MsgBox "No se puede Registrar Scrap Fechas pasados", vbInformation, MSG
'                    CNN.rscmdTurnoScrap.Close
'                     Exit Sub
'                End If
'
'
'                If Turno2 <> Turno Then
'                    MsgBox "No se puede Registrar Scrap Turnos pasados", vbInformation, MSG
'                    CNN.rscmdTurnoScrap.Close
'                      Exit Sub
'                End If
'
'            Else
'
'            End If
'        CNN.rscmdTurnoScrap.Close
'
'
'        Else
'
'        End If
'       ' CNN.rscmdTurnoScrap.Close
'    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim FechaAux As String
FechaAux = Me.LblFecha.Caption

CNN.CmdBuscaIdOperErrorDefX (TipoScrap), (CodScrap)
    If CNN.rsCmdBuscaIdOperErrorDefX.EOF <> True Then
        IdOperError = CNN.rsCmdBuscaIdOperErrorDefX!IdOperError
    Else
        IdOperError = 0
    End If
CNN.rsCmdBuscaIdOperErrorDefX.Close

Dim auxJC As Long
auxJC = CLng(IdJC)
    ''''''''''''''''''''''''''''''''''''''''''' buscar conceptos y operaciones repetidas'''''''''''''''''''''''''''''''''''''''''''''''
    CNN.CmdBuscaScrapRepetido (IdJC), (CodLinea), (TxtClaveP.Text), (IdOperError), (LblNodeParte.Caption), (Etiqueta), (Turno), (Fecha), ("A"), (0), ("No")
    If CNN.rsCmdBuscaScrapRepetido.EOF <> True Then
        MsgBox "Ya existe esta combinacion de Scrap y concepto, necesita eliminarlo para volver a registrar", vbInformation, MSG
        CNN.rsCmdBuscaScrapRepetido.Close
        Exit Sub
    Else
        CNN.rsCmdBuscaScrapRepetido.Close
    End If

    ''''''''''''''''''''''''''''''''''
CodVidrio = Me.LblCodVidrio.Caption
CNN.CmdBuscaDatosJC_VIdrio (CodVidrio)
    ''''CNN.CmdBuscaDatosJC_VIdrio (Fecha), (Turno), (CodLinea), (IdJC) '''''' antes 300821
    If CNN.rsCmdBuscaDatosJC_VIdrio.EOF <> True Then
        
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '' (codlinea = ?) AND (sc_fecha = ?) AND (codturno = ?) AND (Oma_Id = ?)
        CNN.cmdcapscrap (CodLinea), (Fecha), (Turno), (IdJC), (TipoScrap_Scrap), (Etiqueta), ("A"), (0), ("No")
         If CNN.rscmdCapScrap.EOF = True Then
            CNN.rsCmdMaxFolScrap.Open
            If CNN.rsCmdMaxFolScrap.EOF <> True Then
                If IsNumeric(CNN.rsCmdMaxFolScrap!m) Then
                    IdScrap = CNN.rsCmdMaxFolScrap!m + 1
                Else
                    IdScrap = 1
                End If
            End If
            CNN.rsCmdMaxFolScrap.Close
            
            '''''''''''''Buscarl el maximo reg_imp'''''''''''''''''''''''''''''''
            reg_impresion = 0
            CNN.cmdcapscrap_imp (CodLinea), (Fecha), (Turno), (IdJC), (TipoScrap_Scrap), (Etiqueta), ("A")
             If CNN.rscmdCapScrap_imp.EOF <> True Then
                CNN.CmdMaxRegImp (CodLinea), (Fecha), (Turno), (IdJC), (TipoScrap_Scrap), (Etiqueta), ("A")
                If CNN.rsCmdMaxRegImp.EOF <> True Then
                       If IsNumeric(CNN.rsCmdMaxRegImp!m) Then
                           reg_impresion = CNN.rsCmdMaxRegImp!m + 1
                       Else
                           reg_impresion = 1
                       End If
                   End If
                   CNN.rsCmdMaxRegImp.Close
             Else
                reg_impresion = 1
             End If
            CNN.rscmdCapScrap_imp.Close
'            If bandera_SupScrap = False Then
'                FrmValidaUsuario.Show 1
'
'                If bandera_usuario = 0 Then
'                    MsgBox "No podra continuar con la captura hasta que ingrese sus credenciales", vbInformation, MSG
'                    Exit Sub
'                End If
'            ElseIf (FechaProduccion) >= (dia) Then
'                 FrmValidaUsuario.Show 1
'
'                If bandera_usuario = 0 Then
'                    MsgBox "No podra continuar con la captura hasta que ingrese sus credenciales", vbInformation, MSG
'                    Exit Sub
'                End If
'            ' elseif bandera_SupScrap=True and (Fecha) >= (dia)
'            End If
            
            CNN.rscmdCapScrap.AddNew
            CNN.rscmdCapScrap!sc_folio = IdScrap
           If bandera_SupScrap = False Then
               Call BuscaSupervisor
           ElseIf (FechaProduccion) >= (dia) Then
               Call BuscaSupervisor
           Else
                 If CodSupervisor = "" Then
                     CodSupervisor = EmpleadoId
                End If
                
                End If
            
            CNN.rscmdCapScrap!sc_supervisor = CodSupervisor
            End If

        IdScrap = CNN.rscmdCapScrap!sc_folio
        CNN.rscmdCapScrap!codturno = Turno
        
        If bandera_SupScrap = False Then
             CNN.rscmdCapScrap!sc_fecha = Fecha
        Else
            If Time >= CDate("12:00:00 a.m.") And Time <= CDate("07:15:00 a.m.") Then
                CNN.rscmdCapScrap!sc_fecha = Date - 1
            Else
                CNN.rscmdCapScrap!sc_fecha = Date
            End If
        End If
        
        CNN.rscmdCapScrap!CodLinea = CodLinea
        CNN.rscmdCapScrap!reg_imp = 0 ' reg_impresion
        CNN.rscmdCapScrap!impreso = "No"
        
        CNN.rscmdCapScrap!CodVidrio = CNN.rsCmdBuscaDatosJC_VIdrio!CodVidrio
        CNN.rscmdCapScrap!sc_pzas = "F"
        
        If banderaTemplado = 1 Then
            CNN.rscmdCapScrap!sc_templado = "T"
        Else
            CNN.rscmdCapScrap!sc_templado = "ST"
        End If
        
        CNN.rscmdCapScrap!sc_tipo = Mid(CNN.rsCmdBuscaDatosJC_VIdrio!CodVidrio, 7, 3)
        CNN.rscmdCapScrap!sc_espesor = Mid(CNN.rsCmdBuscaDatosJC_VIdrio!CodVidrio, 1, 3)
        CNN.rscmdCapScrap!sc_desorille = "NO"
        CNN.rscmdCapScrap!sc_status = "A"
        CNN.rscmdCapScrap!sc_obs = "° Conteos "

        If IsNull(CNN.rscmdCapScrap!sc_kgs) Then
            CNN.rscmdCapScrap!sc_kgs = (CNN.rsCmdBuscaDatosJC_VIdrio!FactorPeso * Me.TxtCantidad.Text)
        Else
            CNN.rscmdCapScrap!sc_kgs = CNN.rscmdCapScrap!sc_kgs + (CNN.rsCmdBuscaDatosJC_VIdrio!FactorPeso * Me.TxtCantidad.Text)
        End If


'        If TipoScrap_Scrap = "Proveedor" Then
'            CNN.rscmdCapScrap!sc_etiqueta = Etiqueta
'        Else
'            CNN.rscmdCapScrap!sc_etiqueta = CNN.rsCmdBuscaDatosJC_VIdrio!TicketProv
'        End If
        CNN.rscmdCapScrap!sc_etiqueta = Etiqueta
        'CNN.rscmdCapScrap!CodProveedor = CNN.rsCmdBuscaDatosJC_VIdrio!CodProv

        If IsNull(CNN.rscmdCapScrap!sc_cantpza) Then
            CNN.rscmdCapScrap!sc_cantpza = Me.TxtCantidad.Text
        Else
            CNN.rscmdCapScrap!sc_cantpza = CNN.rscmdCapScrap!sc_cantpza + Me.TxtCantidad.Text
        End If
        CNN.rscmdCapScrap!sc_factor = CNN.rsCmdBuscaDatosJC_VIdrio!FactorPeso
        If IsNull(CNN.rscmdCapScrap!sc_pesocal) Then
            CNN.rscmdCapScrap!sc_pesocal = (CNN.rsCmdBuscaDatosJC_VIdrio!FactorPeso * Me.TxtCantidad.Text)
        Else
            CNN.rscmdCapScrap!sc_pesocal = CNN.rscmdCapScrap!sc_pesocal + (CNN.rsCmdBuscaDatosJC_VIdrio!FactorPeso * Me.TxtCantidad.Text)
        End If
        CNN.rscmdCapScrap!sc_tmerma = ""
        CNN.rscmdCapScrap!sc_tscrap = TipoScrap_Scrap
        CNN.rscmdCapScrap!sc_dif = 0
        CNN.rscmdCapScrap!sc_ptara = 0
        If IsNull(CNN.rscmdCapScrap!sc_pneto) Then
            CNN.rscmdCapScrap!sc_pneto = (CNN.rsCmdBuscaDatosJC_VIdrio!FactorPeso * Me.TxtCantidad.Text)
        Else
            CNN.rscmdCapScrap!sc_pneto = CNN.rscmdCapScrap!sc_pneto + (CNN.rsCmdBuscaDatosJC_VIdrio!FactorPeso * Me.TxtCantidad.Text)
        End If
        
'''     WHERE (dbo.MatrizMPS.codigo = ?) AND (dbo.Tbl_DetOma.Oma_Id = ?) AND
'''    (dbo.Tbl_DetOma.FechaCap = ?) AND (dbo.TblHorariosHxH.Turno = ?)
'''    AND (dbo.Tbl_DetOma.codlinea = ?)
'        Dim a
'        a = CNN.rsCmdBuscaDatosJC_VIdrio!OF
'        If CNN.rsCmdBuscaDatosJC_VIdrio!Tipo = "AT" Then
'            CNN.CmdBuscaDatosJC_VIdrioAT (CNN.rsCmdBuscaDatosJC_VIdrio!OF), (IdJC), (Fecha), (Turno), (CodLinea)
'            If CNN.rsCmdBuscaDatosJC_VIdrioAT.EOF <> True Then
'                OF = CNN.rsCmdBuscaDatosJC_VIdrioAT!OF
'                CNN.rscmdCapScrap!OF = CNN.rsCmdBuscaDatosJC_VIdrioAT!OF
'            End If
'            CNN.rsCmdBuscaDatosJC_VIdrioAT.Close
'        Else
'             OF = CNN.rsCmdBuscaDatosJC_VIdrio!OF
'             CNN.rscmdCapScrap!OF = CNN.rsCmdBuscaDatosJC_VIdrio!OF
'        End If
        CNN.rscmdCapScrap!sc_tscrapmps = 0
        CNN.rscmdCapScrap!sc_nopartemps = ""
       
        If bandera_SupScrap = False Then
            CNN.rscmdCapScrap!EmpleadoId = EmpleadoId
        Else
            'buscar empleado que estuvo trabajando ese dia de fabricacion'
            CNN.rscmdCapScrap!EmpleadoId = EmpleadoId
'            If IdWip <> "" Then
'                CNN.rscmdCapScrap!EmpleadoId = EmpleadoWip
'                CNN.rscmdCapScrap!IdAcum = IdWip
'                CNN.rscmdCapScrap!sc_obs = "° Conteos, El folio de wip es:" & IdWip & ", con Fecha de produccion:" & FechaProduccion
'            End If
            
        End If
        CNN.rscmdCapScrap!sc_cargoa = ""
        CNN.rscmdCapScrap!Oma_Id = IdJC
        CNN.rscmdCapScrap!sc_fecha_produccion = Fecha 'FechaProduccion
        CNN.rscmdCapScrap!codturno_produccion = LblTurno.Caption
        CNN.rscmdCapScrap!EmpleadoID_registra = EmpleadoId
        
        
        CNN.rscmdCapScrap.Update
        CNN.rscmdCapScrap.Close

        CNN.CmdBuscaIdOperError (IdOperacion), (CodScrap)
        If CNN.rsCmdBuscaIdOperError.EOF <> True Then
            IdOperError = CNN.rsCmdBuscaIdOperError!IdOperError
        Else
            IdOperError = 0
        End If
        CNN.rsCmdBuscaIdOperError.Close

        Call MaxDetScrap

        CNN.CmdDetScrapCont (IdScrap), (DetScrap)
        If CNN.rsCmdDetScrapCont.EOF = True Then
            CNN.rsCmdDetScrapCont.AddNew
            CNN.rsCmdDetScrapCont!sc_folio = IdScrap
            CNN.rsCmdDetScrapCont!coderror = CodScrap
            CNN.rsCmdDetScrapCont!scd_item = DetScrap
            CNN.rsCmdDetScrapCont!scd_cantidad = Me.TxtCantidad.Text
            CNN.rsCmdDetScrapCont!Proceso = IdOperacion    'vProceso 'clave de proceso
            'CNN.rsCmdDetScrapCont!OF = CNN.rsCmdBuscaDatosJC_VIdrio!OF
            CNN.rsCmdDetScrapCont!OF = OF
            CNN.rsCmdDetScrapCont!scd_Obs = " ° Conteos Det  "
            CNN.rsCmdDetScrapCont!scd_kilos = (CNN.rsCmdBuscaDatosJC_VIdrio!FactorPeso * Me.TxtCantidad.Text)         'kilos
            CNN.rsCmdDetScrapCont!IdOperError = IdOperError
            CNN.rsCmdDetScrapCont.Update
        End If
        CNN.rsCmdDetScrapCont.Close
    Else

    End If
    CNN.rsCmdBuscaDatosJC_VIdrio.Close

    '''''''''''''''''''''''''''''dimensiones'''''''''''''''''''''''''''''''''''
'
'    If Me.CboTipoSrap.Text = "PROVEEDOR" Then
'        TicketProv = ""
'        'FrmLstEtiquetasMP.Show 1
'
'
'        '    If TicketProv = "SALIR" Or TicketProv = "" Then
'        '
'        '          MsgBox "No Selcciono una etiqueta.", vbInformation, MSG
'        '          Exit Sub
'        '     Else
'        If Me.CboConceptoScrap.Text = "DIMENSION" Then
'            DetScrap = DetScrap
'            IdScrap = IdScrap
'            CodScrap = CodScrap
'            FrmDimensiones.Show 1
'        End If
'        '     End If
'        TipoScrap_Scrap = "Proveedor"
'        If banderaTemplado = 1 Then
'            TipoScrap_Scrap = "Templado"
'         Else
'              TipoScrap_Scrap = "Proceso"
'         End If
''         'If CodLinea = 152 Then
''           TipoScrap_Scrap = "Retrabajo"
''        'End If
'    Else
'        If banderaTemplado = 1 Then
'            TipoScrap_Scrap = "Templado"
'         Else
'              TipoScrap_Scrap = "Proceso"
'              If Bandera_Retrabajo = "Retrabajo" Then
'                    TipoScrap_Scrap = "Retrabajo"
'              End If
'
'         End If
''         'If CodLinea = 152 Then
''           TipoScrap_Scrap = "Retrabajo"
''        'End If
'    End If


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MsgBox "Se guardo Correctamente", vbInformation, MSG

  

'    Call LlenaTipoScrap
'  Call ConceptosScrap
'Carga Scrap Capturado para la operacion
'  Call LlenarDetsScrap
'  Call CargaLstScrap
'Call LlenarTotalesxTipo
    
'Detalle = 1
''CmdDetScrapContadores
'CNN.CmdDetScrapContadores (IdJC), (IdOperacion), (CodScrap), (Turno), (Fecha)
'If CNN.rsCmdDetScrapContadores.EOF = True Then
'            CNN.rsCmdDetScrapContadores.AddNew
'End If
'            'CNN.rsCmdDetScrapContadores!IdDetScrap = Detalle
'            CNN.rsCmdDetScrapContadores!Item = Detalle
'            CNN.rsCmdDetScrapContadores!Oma_Id = IdJC
'            CNN.rsCmdDetScrapContadores!codopera = IdOperacion
'            CNN.rsCmdDetScrapContadores!coderror = Me.TxtClaveP.Text
'            CNN.rsCmdDetScrapContadores!Cantidadpzs = Me.TxtCantidad.Text
'            ' 3= 8
'            ' 4= 10
'            ' 5= 12.485
'            ' 6= 14.99
'            ' 8= 16.88
''                CNN.CmdBuscaVidrioKgs (CodVidrio)
''                If CNN.rsCmdBuscaVidrioKgs.EOF <> True Then
''                Else
''                End If
''                CNN.rsCmdBuscaVidrioKgs.Close
'            CNN.rsCmdDetScrapContadores!CantidadKg = 0
'            CNN.rsCmdDetScrapContadores!Cantidadm2 = 0
'            CNN.rsCmdDetScrapContadores!Status = "A"
'            CNN.rsCmdDetScrapContadores!Observacioens = ""
'            CNN.rsCmdDetScrapContadores!EmpleadoId = EmpleadoId
'            CNN.rsCmdDetScrapContadores!fechacap = Fecha          'Date
'            CNN.rsCmdDetScrapContadores!horacap = Time
'            CNN.rsCmdDetScrapContadores!Turno = Turno
'            CNN.rsCmdDetScrapContadores!Cantidadpzssup = Me.TxtCantidad.Text
'    CNN.rsCmdDetScrapContadores.Update
'
'    Detalle = CNN.rsCmdDetScrapContadores!IdDetScrap
'
'    ''Actualiza JC
'    CNN.CmdBuscaScrapCapturadoxOperacion (IdOperacion), (IdJC), (Turno)
'    If CNN.rsCmdBuscaScrapCapturadoxOperacion.EOF <> True Then
'        j = 1
'        Do While j <= (Renglon - 1)
'            If Me.LstTM.ListItems(j).ListSubItems.Item(13).Text = "<---" Then
'                    Me.LstTM.ListItems.Item(j).SubItems(10) = CNN.rsCmdBuscaScrapCapturadoxOperacion!TotalScrap
'                    TScrapxOperacion = CNN.rsCmdBuscaScrapCapturadoxOperacion!TotalScrap
'                    Exit Do
'            End If
'            j = j + 1
'        Loop
'    Else
'        Me.LstTM.ListItems.Item(RenglonSeleccionado).SubItems(10) = 0
'    End If
'    CNN.rsCmdBuscaScrapCapturadoxOperacion.Close
'
''     If (TScrapxOperacion) > (TotalScrapxCapturar) Then
''        MsgBox "No puede agregar esta cantidad de scrap, superaria el total de Scrap de la operacion.", vbCritical, MSG
''
''        CNN.CmdDeleteDetSCrapRenglonDets (Detalle)
''        If CNN.rsCmdDeleteDetScrapRenglonDets.EOF <> True Then
''            CNN.rsCmdDeleteDetScrapRenglonDets.Delete
''            CNN.rsCmdDeleteDetScrapRenglonDets.Update
''        End If
''        CNN.rsCmdDeleteDetScrapRenglonDets.Close
''
''    '''    Exit Sub
''
''        ''Actualiza JC
''        CNN.CmdBuscaScrapCapturadoxOperacion (IdOperacion), (IdJC), (Turno)
''        If CNN.rsCmdBuscaScrapCapturadoxOperacion.EOF <> True Then
''            j = 1
''            Do While j <= (Renglon - 1)
''                If Me.LstTM.ListItems(j).ListSubItems.Item(13).Text = "<---" Then
''                        Me.LstTM.ListItems.Item(j).SubItems(10) = CNN.rsCmdBuscaScrapCapturadoxOperacion!TotalScrap
''                        Exit Do
''                End If
''                j = j + 1
''            Loop
''        Else
''            Me.LstTM.ListItems.Item(RenglonSeleccionado).SubItems(10) = 0
''        End If
''        CNN.rsCmdBuscaScrapCapturadoxOperacion.Close
''    Else
''        MsgBox "Se guardo Correctamente", vbInformation, MSG
''    End If
'
'    MsgBox "Se guardo Correctamente", vbInformation, MSG
'
'CNN.rsCmdDetScrapContadores.Close
'
'''Carga Scrap Capturado para la operacion
Call LlenarDetsScrap

End Sub


