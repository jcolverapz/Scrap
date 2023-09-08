Public Sub LlenarDetsScrap() ' capturados
    Fecha = Me.labelFecha.Caption
    Turno = Me.labelTurno.Caption
    Etiqueta = LabelEtiquetaWip.Caption
    TScrapxOperacion = 0
    Me.LstDetsScrap.ListItems.Clear
    RenglonDets = 1
    
    If bandera_SupScrap = False Then
             Fecha = Fecha
    Else
        If (FechaProduccion) >= (dia) Then
            If Time >= CDate("12:00:00 a.m.") And Time <= CDate("07:15:00 a.m.") Then
               Fecha = Date - 1
            Else
                Fecha = Date
            End If
        End If
        
    End If
    
    CNN.CmdLstScrapDetEtiquetaWip (Fecha), (CodLinea), (Turno), (IdJC), (TicketProv), (Etiqueta), (0), ("No") ', (IdOperacion)
    Do While CNN.rsCmdLstScrapDetEtiquetaWip.EOF <> True
        Me.LstDetsScrap.ListItems.Add (RenglonDets)
        'Llena Etiquetas
        Me.LstDetsScrap.ListItems.Item(RenglonDets).SubItems(1) = CNN.rsCmdLstScrapDetEtiquetaWip!Descripcion
        Me.LstDetsScrap.ListItems.Item(RenglonDets).SubItems(2) = CNN.rsCmdLstScrapDetEtiquetaWip!Defecto
        Me.LstDetsScrap.ListItems.Item(RenglonDets).SubItems(3) = CNN.rsCmdLstScrapDetEtiquetaWip!pzs
        '        Me.LstDetsScrap.ListItems.Item(RenglonDets).SubItems(4) = CNN.rsCmdLstScrapDetEtiqueta!scd_item
        Me.LstDetsScrap.ListItems.Item(RenglonDets).SubItems(5) = ""
        Me.LstDetsScrap.ListItems.Item(RenglonDets).SubItems(6) = CNN.rsCmdLstScrapDetEtiquetaWip!sc_folio
        Me.LstDetsScrap.ListItems.Item(RenglonDets).SubItems(7) = CNN.rsCmdLstScrapDetEtiquetaWip!scd_item

        Me.LstDetsScrap.ListItems.Item(RenglonDets).SubItems(8) = CNN.rsCmdLstScrapDetEtiquetaWip!sc_tscrap
        Me.LstDetsScrap.ListItems.Item(RenglonDets).SubItems(10) = CNN.rsCmdLstScrapDetEtiquetaWip!sc_etiqueta
        IdScrap = CNN.rsCmdLstScrapDetEtiquetaWip!sc_folio
        TScrapxOperacion = TScrapxOperacion + CNN.rsCmdLstScrapDetEtiquetaWip!pzs

        RenglonDets = RenglonDets + 1
        CNN.rsCmdLstScrapDetEtiquetaWip.MoveNext
    Loop
    CNN.rsCmdLstScrapDetEtiquetaWip.Close

    LblCapturado.Caption = "   Total Capturado x Operacion:  " & TScrapxOperacion & "  "

End Sub

Public Sub LlenarDetsScrapxOperacion() ' capturados
    Fecha = Me.labelFecha.Caption
    Etiqueta = LabelEtiquetaWip.Caption
    TScrapxOperacion = 0
    Me.LstDetsScrap.ListItems.Clear
    RenglonDets = 1
    
    If bandera_SupScrap = False Then
             Fecha = Fecha
    Else
     If (FechaProduccion) >= (dia) Then
        If Time >= CDate("12:00:00 a.m.") And Time <= CDate("07:15:00 a.m.") Then
           Fecha = Date - 1
        Else
            Fecha = Date
        End If
      End If
      
    End If
    
    CNN.CmdBuscaxOperacionWip (Fecha), (CodLinea), (Turno), (IdJC), (Me.CboTipoSrap.Text), (TicketProv), ("P"), (Etiqueta), (0), ("No")
    Do While CNN.rsCmdBuscaxOperacionWip.EOF <> True
        Me.LstDetsScrap.ListItems.Add (RenglonDets)
        'Llena Etiquetas
        Me.LstDetsScrap.ListItems.Item(RenglonDets).SubItems(1) = CNN.rsCmdBuscaxOperacionWip!Descripcion
        Me.LstDetsScrap.ListItems.Item(RenglonDets).SubItems(2) = CNN.rsCmdBuscaxOperacionWip!Defecto
        Me.LstDetsScrap.ListItems.Item(RenglonDets).SubItems(3) = CNN.rsCmdBuscaxOperacionWip!pzs
        '        Me.LstDetsScrap.ListItems.Item(RenglonDets).SubItems(4) = CNN.rsCmdLstScrapDetEtiqueta!scd_item
        Me.LstDetsScrap.ListItems.Item(RenglonDets).SubItems(5) = ""
        Me.LstDetsScrap.ListItems.Item(RenglonDets).SubItems(6) = CNN.rsCmdBuscaxOperacionWip!sc_folio
        Me.LstDetsScrap.ListItems.Item(RenglonDets).SubItems(7) = CNN.rsCmdBuscaxOperacionWip!scd_item

        Me.LstDetsScrap.ListItems.Item(RenglonDets).SubItems(8) = CNN.rsCmdBuscaxOperacionWip!sc_tscrap
        Me.LstDetsScrap.ListItems.Item(RenglonDets).SubItems(10) = CNN.rsCmdBuscaxOperacionWip!sc_etiqueta

        TScrapxOperacion = TScrapxOperacion + CNN.rsCmdBuscaxOperacionWip!pzs

        RenglonDets = RenglonDets + 1
        CNN.rsCmdBuscaxOperacionWip.MoveNext
    Loop
    CNN.rsCmdBuscaxOperacionWip.Close

    LblCapturado.Caption = "   Total Capturado x Operacion:  " & TScrapxOperacion & "  "

End Sub