Private Sub CmdImprimir_Click()
    If bandera_SupScrap = False Then
             Fecha = Fecha
        Else
            If Time >= CDate("12:00:00 a.m.") And Time <= CDate("07:15:00 a.m.") Then
                Fecha = Date - 1
            Else
                Fecha = Date
            End If
    End If
        
    If TScrapxOperacion <> 0 Then
    Dim idFolio
    idFolio = ""
'''''      CNN.CmdLstScrapDetEtiqueta (Fecha), (CodLinea), (Turno), (IdJC), (Etiqueta), (TipoScrap_Scrap) ', (IdOperacion)
'''''        Do While CNN.rsCmdLstScrapDetEtiqueta.EOF <> True
'''''            IdScrap = CNN.rsCmdLstScrapDetEtiqueta!sc_folio
'''''            If idFolio <> IdScrap Then
'''''                ' FrmRepScrapEtaRoja.Show 1
'''''               ' FrmRepScrapEtaRojaProve.Show 1
'''''            End If
'''''            idFolio = IdScrap
'''''            CNN.rsCmdLstScrapDetEtiqueta.MoveNext
'''''        Loop
'''''        CNN.rsCmdLstScrapDetEtiqueta.Close
''''
''''         'FrmRepScrapEtaRojaProve.Show 1
''''         FrmRepScrapEtaRojaProve.Show 1
''''        'Form1.Show 1


        
        CNN.CmdLstScrapDetEtiqueta (Fecha), (CodLinea), (Turno), (IdJC), (Etiqueta), (TipoScrap_Scrap), (0), ("No")   ', (IdOperacion)
       ' CNN.CmdLstScrapDetEtiqueta (Fecha), (CodLinea), (Turno), (IdJC), (Etiqueta), ("Retrabajo"), (0), ("No")   ', (IdOperacion)
        If CNN.rsCmdLstScrapDetEtiqueta.EOF <> True Then
           IdScrap = CNN.rsCmdLstScrapDetEtiqueta!sc_folio
             If banderaTemplado <> 1 Then
                     CNN.CmdImpresoBoton (IdScrap)
                     If CNN.rsCmdImpresoBoton.EOF <> True Then
                        If CNN.rsCmdImpresoBoton!reg_imp = 0 Then
                             CNN.rsCmdImpresoBoton.Close
                             FrmRepScrapEtaRoja.Show 1
                        Else
                           MsgBox ("Este folio ya fue impreso " & IdScrap & "Con fecha de impresión" & CNN.rsCmdImpresoBoton!Fecha_imp)
                           CNN.rsCmdImpresoBoton.Close
                        End If
                        
                     End If
    '                If CNN.rsCmdImpresoBoton.State = 1 Then
    '                    CNN.rsCmdImpresoBoton.Close
    '                End If
               End If
            
      
         CNN.rsCmdLstScrapDetEtiqueta.Close
        Else
            CNN.rsCmdLstScrapDetEtiqueta.Close
        End If
        
        If banderaTemplado = 1 Then
'            CNN.CmdLstScrapDetEtiqueta (Fecha), (CodLinea), (Turno), (IdJC), (Etiqueta), ("Templado"), (0), ("No")   ', (IdOperacion)
'            Do While CNN.rsCmdLstScrapDetEtiqueta.EOF <> True
'                IdScrap = CNN.rsCmdLstScrapDetEtiqueta!sc_folio
'                FrmRepScrapEtaRoja.Show 1
'                Exit Do
'            Loop
'            CNN.rsCmdLstScrapDetEtiqueta.Close
            
             
        '''''''''''''''''Imprimir cerrar Scrap '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'            CNN.cmdcapscrap (CodLinea), (Fecha), (Turno), (IdJC), (TipoScrap_Scrap), (Etiqueta), ("P"), (0), ("No")
'             If CNN.rscmdCapScrap.EOF <> True Then
'                CNN.rscmdCapScrap!impreso = "Si"
'                CNN.rscmdCapScrap.Update
'             End If
'            CNN.rscmdCapScrap.Close
'            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'             Unload Me
                CNN.CmdImpresoBoton (IdScrap)
                 If CNN.rsCmdImpresoBoton.EOF <> True Then
                    If CNN.rsCmdImpresoBoton!reg_imp = 0 Then
                         CNN.rsCmdImpresoBoton.Close
                         FrmRepScrapEtaRoja.Show 1
                    Else
                        MsgBox ("Este folio ya fue impreso " & IdScrap)
                    End If
                    
                 End If
          Unload Me
        End If
    '''''''''''Verificar si contien proveedores dentro del TipoScrap Proceso'''''''''''''''''''''''
     If banderaTemplado <> 1 Then
     
       CNN.CmdLstScrapDetEtiqueta (Fecha), (CodLinea), (Turno), (IdJC), (Etiqueta), ("Proveedor"), (0), ("No")  ', (IdOperacion)
       
        If CNN.rsCmdLstScrapDetEtiqueta.EOF <> True Then
            IdScrap = CNN.rsCmdLstScrapDetEtiqueta!sc_folio
            '''''''''''''''''''Verificar si existe la combinacion de proveedor dimension'''''''''''''''''''''''
             CNN.CmdBuscaDim (IdScrap)
            If CNN.rsCmdBuscaDim.EOF <> True Then
              Call numeroDimensiones
               If dimension_item = 10 Then
                 CNN.rsCmdBuscaDim.Close
                 FrmRepScrapEtaRojaProve.Show 1
                Else
                    CNN.rsCmdBuscaDim.Close
                    MsgBox "Complete las 10 dimensiones para imprimir el reporte de Proveedor"
                End If
            Else
                CNN.rsCmdBuscaDim.Close
                CNN.CmdImpresoBoton (IdScrap)
                 If CNN.rsCmdImpresoBoton.EOF <> True Then
                    If CNN.rsCmdImpresoBoton!reg_imp = 0 Then
                         CNN.rsCmdImpresoBoton.Close
                         FrmRepScrapEtaRojaProve.Show 1
                    Else
                        MsgBox ("Este folio ya fue impreso " & IdScrap)
                    End If
                    
                 End If
'                If CNN.rsCmdImpresoBoton.State = 1 Then
'                    CNN.rsCmdImpresoBoton.Close
'                End If
                
            End If
            
'            Exit Do
'        Loop
        End If
        CNN.rsCmdLstScrapDetEtiqueta.Close
    
    
    '''''''''''''''''Imprimir cerrar Scrap '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''
''''        CNN.cmdcapscrap (CodLinea), (Fecha), (Turno), (IdJC), (TipoScrap_Scrap), (Etiqueta), ("P"), (0), ("No")
''''         If CNN.rscmdCapScrap.EOF <> True Then
''''            CNN.rscmdCapScrap!impreso = "Si"
''''            CNN.rscmdCapScrap.Update
''''         End If
''''        CNN.rscmdCapScrap.Close
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         Unload Me
    End If
     Else
        MsgBox "No existe Scrap para imprimir.", vbInformation, MSG
    End If
End Sub