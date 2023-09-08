SELECT TOP 300 Cap_scrap.sc_folio, Cap_scrap.codturno,
    Cap_scrap.sc_fecha, lineas.descripcion AS Linea, Orden_Man.Oma_Id,
    Orden_Man.Oma_Tipo, Det_PackList.TicketProv, Tbl_Partes.Parte,
    Vidrio.CodVidrio, Vidrio.Codigo_Interno, Vidrio.ESPESOR, Vidrio.Color,
    Vidrio.X, Vidrio.Y, Vidrio.Cstd_Vidrio, Clientes.CodCliente,
    Clientes.NombreCorto, Cap_scrap.sc_supervisor,
    EMPLEADO1.EMP_NOMBRE AS Capturo, Cap_scrap.sc_templado,
    Cap_scrap.sc_pzas, Cap_scrap.sc_tipo, Cap_scrap.sc_desorille,
    Cap_scrap.sc_status, Cap_scrap.sc_obs, Cap_scrap.sc_kgs,
    Proveedores.CodProveedor, Proveedores.Nom_Corto,
    Cap_scrap.sc_factor, Cap_scrap.sc_pesocal, Cap_scrap.sc_tscrap,
    Cap_scrap.sc_pneto, Cap_scrap.EmpleadoID,
    EMPLEADO1.EMP_NOMBRE AS Capturo, Cap_scrap.reg_imp,
    Cap_scrap.impreso, operaciones.codopera,
    operaciones.descripcion AS Operacion, det_scrap.scd_item,
    det_scrap.coderror, Defectos.descripcion AS Defecto,
    det_scrap.scd_cantidad, det_scrap.scd_obs, det_scrap.scd_kilos,
    Tbl_Partes.X AS X_PT, Tbl_Partes.Y AS Y_PT, Tbl_Partes.FactorPeso,
    Tbl_Partes.Comodity, Tbl_Partes.Sap_cc, Tbl_Partes.TipoPrograma,
    Tbl_Partes.TipoSubPrograma, Tbl_Partes.Descripcion,
    EMPLEADO1.EMP_NOMBRE, Det_PackList.TicketGem,
    EMPLEADO2.EMP_NOMBRE AS Supervisor
FROM Clientes INNER JOIN
    det_scrap INNER JOIN
    Cap_scrap ON det_scrap.sc_folio = Cap_scrap.sc_folio INNER JOIN
    DefectosXOperacion ON
    det_scrap.IdOperError = DefectosXOperacion.IdOperError INNER JOIN
    operaciones ON
    DefectosXOperacion.codopera = operaciones.codopera INNER JOIN
    Defectos ON
    DefectosXOperacion.coderror = Defectos.coderror INNER JOIN
    lineas ON Cap_scrap.codlinea = lineas.codlinea INNER JOIN
    Orden_Man ON
    Cap_scrap.Oma_Id = Orden_Man.Oma_Id INNER JOIN
    Tbl_Partes ON Orden_Man."OF" = Tbl_Partes."OF" ON
    Clientes.CodCliente = Tbl_Partes.CodCliente INNER JOIN
    EMPLEADO EMPLEADO1 ON
    Cap_scrap.EmpleadoID = EMPLEADO1.EMPLEADOID INNER JOIN
    Proveedores INNER JOIN
    Vidrio INNER JOIN
    Det_PackList ON Vidrio.CodVidrio = Det_PackList.Codigo ON
    Proveedores.CodProveedor = Vidrio.Codproveedor ON
    Cap_scrap.sc_Etiqueta = Det_PackList.TicketGem INNER JOIN
    EMPLEADO EMPLEADO2 ON
    Cap_scrap.sc_supervisor = EMPLEADO2.EMPLEADOID
ORDER BY Cap_scrap.sc_folio DESC