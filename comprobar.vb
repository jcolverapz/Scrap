'comprobar 

SELECT dbo.Cap_scrap.sc_folio, dbo.Tbl_Partes.Parte,
    dbo.Cap_scrap.sc_Etiqueta, dbo.Cap_scrap.sc_fecha,
    dbo.EMPLEADO.EMP_NOMBRE AS Capturo,
    EMPLEADO1.EMP_NOMBRE AS Supervisor, dbo.Cap_scrap.Oma_Id,
    dbo.Cap_scrap.sc_pzas, dbo.Clientes.NombreCorto,
    dbo.Cap_scrap.codturno, dbo.lineas.descripcion AS Linea,
    dbo.operaciones.descripcion AS Operacion,
    dbo.Defectos.descripcion AS Defecto, dbo.Cap_scrap.sc_tscrap,
    dbo.Cap_scrap.codvidrio, dbo.Cap_scrap.sc_status,
    dbo.Cap_scrap.sc_obs, dbo.Cap_scrap.sc_kgs,
    dbo.det_scrap.scd_cantidad, dbo.Cap_scrap.sc_cantpza,
    dbo.Cap_scrap.sc_pesocal, dbo.Cap_scrap.sc_factor,
    dbo.Cap_scrap.[of], dbo.Proveedores.Nom_Corto,
    dbo.det_scrap.scd_item, dbo.det_scrap.scd_obs,
    dbo.det_scrap.proceso , dbo.det_scrap.scd_kilos
    FROM dbo.Clientes INNER JOIN
    dbo.Tbl_Partes ON
    dbo.Clientes.CodCliente = dbo.Tbl_Partes.CodCliente INNER JOIN
    dbo.det_scrap INNER JOIN
    dbo.Cap_scrap ON
    dbo.det_scrap.sc_folio = dbo.Cap_scrap.sc_folio INNER JOIN
    dbo.Defectos INNER JOIN
    dbo.tipodef ON dbo.Defectos.codtipoe = dbo.tipodef.codtipoe ON
    dbo.det_scrap.coderror = dbo.Defectos.coderror ON
    dbo.Tbl_Partes.[OF] = dbo.Cap_scrap.[of] INNER JOIN
    dbo.EMPLEADO ON
    dbo.Cap_scrap.EmpleadoID = dbo.EMPLEADO.EMPLEADOID INNER JOIN
    dbo.EMPLEADO EMPLEADO1 ON
    dbo.Cap_scrap.sc_supervisor = EMPLEADO1.EMPLEADOID INNER JOIN
    dbo.lineas ON
    dbo.Cap_scrap.codlinea = dbo.lineas.codlinea INNER JOIN
    dbo.DefectosXOperacion ON
    dbo.det_scrap.IdOperError = dbo.DefectosXOperacion.IdOperError INNER JOIN
    dbo.operaciones ON
    dbo.DefectosXOperacion.codopera = dbo.operaciones.codopera INNER JOIN
    dbo.Proveedores ON
    dbo.Cap_scrap.CodProveedor = dbo.Proveedores.CodProveedor
   
