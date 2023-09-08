'si funciona
SELECT        Cap_scrap.sc_folio, Tbl_Partes.Parte, Cap_scrap.sc_Etiqueta, Cap_scrap.sc_fecha, EMPLEADO.EMP_NOMBRE AS Capturo, EMPLEADO1.EMP_NOMBRE AS Supervisor, Cap_scrap.Oma_Id, Cap_scrap.sc_pzas, 
                         Clientes.NombreCorto, Cap_scrap.codturno, lineas.descripcion AS Linea, operaciones.descripcion AS Operacion, Defectos.descripcion AS Defecto, Cap_scrap.sc_tscrap, Cap_scrap.codvidrio, Cap_scrap.sc_status, 
                         Cap_scrap.sc_obs, Cap_scrap.sc_kgs, det_scrap.scd_cantidad, Cap_scrap.sc_cantpza, Cap_scrap.sc_pesocal, Cap_scrap.sc_factor, Cap_scrap.[of], det_scrap.scd_item, det_scrap.scd_obs, det_scrap.proceso, 
                         det_scrap.scd_kilos
FROM            Clientes INNER JOIN
                         Tbl_Partes ON Clientes.CodCliente = Tbl_Partes.CodCliente INNER JOIN
                         det_scrap INNER JOIN
                         Cap_scrap ON det_scrap.sc_folio = Cap_scrap.sc_folio INNER JOIN
                         Defectos INNER JOIN
                         tipodef ON Defectos.codtipoe = tipodef.codtipoe ON det_scrap.coderror = Defectos.coderror ON Tbl_Partes.[OF] = Cap_scrap.[of] INNER JOIN
                         EMPLEADO ON Cap_scrap.EmpleadoID = EMPLEADO.EMPLEADOID INNER JOIN
                         EMPLEADO AS EMPLEADO1 ON Cap_scrap.sc_supervisor = EMPLEADO1.EMPLEADOID INNER JOIN
                         lineas ON Cap_scrap.codlinea = lineas.codlinea INNER JOIN
                         DefectosXOperacion ON det_scrap.IdOperError = DefectosXOperacion.IdOperError INNER JOIN
                         operaciones ON DefectosXOperacion.codopera = operaciones.codopera