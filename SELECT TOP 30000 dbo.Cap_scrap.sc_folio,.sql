SELECT TOP 30000 dbo.Cap_scrap.sc_folio, dbo.Cap_scrap.codturno, 
    dbo.Cap_scrap.sc_fecha, dbo.lineas.descripcion AS Linea, 
    dbo.Orden_Man.Oma_Id, dbo.Orden_Man.Oma_Tipo, 
    dbo.Det_PackList.TicketProv, dbo.Tbl_Partes.Parte, 
    dbo.Vidrio.CodVidrio, dbo.Vidrio.Codigo_Interno, dbo.Vidrio.ESPESOR, 
    dbo.Vidrio.Color, dbo.Vidrio.X, dbo.Vidrio.Y, dbo.Vidrio.Cstd_Vidrio, 
    dbo.Clientes.CodCliente, dbo.Clientes.NombreCorto, 
    dbo.Cap_scrap.sc_supervisor, 
    EMPLEADO1.EMP_NOMBRE AS Supervisor, 
    dbo.Cap_scrap.sc_templado, dbo.Cap_scrap.sc_pzas, 
    dbo.Cap_scrap.sc_tipo, dbo.Cap_scrap.sc_desorille, 
    dbo.Cap_scrap.sc_status, dbo.Cap_scrap.sc_obs, 
    dbo.Cap_scrap.sc_kgs, dbo.Proveedores.CodProveedor, 
    dbo.Proveedores.Nom_Corto, dbo.Cap_scrap.sc_factor, 
    dbo.Cap_scrap.sc_pesocal, dbo.Cap_scrap.sc_tscrap, 
    dbo.Cap_scrap.sc_pneto, dbo.Cap_scrap.EmpleadoID, 
    EMPLEADO1.EMP_NOMBRE AS Capturo, dbo.Cap_scrap.reg_imp, 
    dbo.Cap_scrap.impreso, dbo.operaciones.codopera, 
    dbo.operaciones.descripcion AS Operacion, dbo.det_scrap.scd_item, 
    dbo.det_scrap.coderror, dbo.Defectos.descripcion AS Defecto, 
    dbo.det_scrap.scd_cantidad, dbo.det_scrap.scd_obs, 
    dbo.det_scrap.scd_kilos, dbo.Tbl_Partes.X AS X_PT, 
    dbo.Tbl_Partes.Y AS Y_PT, dbo.Tbl_Partes.FactorPeso, 
    dbo.Tbl_Partes.Comodity, dbo.Tbl_Partes.Sap_cc, 
    dbo.Tbl_Partes.TipoPrograma, dbo.Tbl_Partes.TipoSubPrograma, 
    dbo.Tbl_Partes.Descripcion, dbo.EMPLEADO.EMP_NOMBRE
FROM dbo.det_scrap INNER JOIN
    dbo.Cap_scrap ON 
    dbo.det_scrap.sc_folio = dbo.Cap_scrap.sc_folio INNER JOIN
    dbo.DefectosXOperacion ON 
    dbo.det_scrap.IdOperError = dbo.DefectosXOperacion.IdOperError 
    INNER JOIN
    dbo.operaciones ON 
    dbo.DefectosXOperacion.codopera = dbo.operaciones.codopera INNER JOIN    dbo.Defectos 
    ON dbo.DefectosXOperacion.coderror = dbo.Defectos.coderror INNER JOIN    dbo.lineas 
    ON dbo.Cap_scrap.codlinea = dbo.lineas.codlinea INNER JOIN dbo.Orden_Man 
    ON dbo.Cap_scrap.Oma_Id = dbo.Orden_Man.Oma_Id INNER JOIN dbo.Det_PackList 
    ON dbo.Cap_scrap.sc_Etiqueta = dbo.Det_PackList.TicketProv INNER JOIN  dbo.Tbl_Partes 
    ON dbo.Orden_Man."OF" = dbo.Tbl_Partes."OF" INNER JOIN    dbo.Vidrio 
    ON dbo.Det_PackList.Codigo = dbo.Vidrio.CodVidrio INNER JOIN   dbo.Clientes 
    ON dbo.Tbl_Partes.CodCliente = dbo.Clientes.CodCliente INNER JOIN  dbo.EMPLEADO EMPLEADO1 
    ON dbo.Cap_scrap.EmpleadoID = EMPLEADO1.EMPLEADOID INNER JOIN   dbo.Proveedores 
    ON dbo.Vidrio.Codproveedor = dbo.Proveedores.CodProveedor INNER JOIN  dbo.EMPLEADO 
    ON dbo.Cap_scrap.sc_supervisor = dbo.EMPLEADO.EMPLEADOID
    
    ORDER BY dbo.Cap_scrap.sc_folio DESC
