SELECT TOP 100 dbo.Cap_scrap.sc_fecha, dbo.Cap_scrap.codlinea, 
    dbo.Cap_scrap.codturno, dbo.Cap_scrap.Oma_Id, 
    dbo.DefectosXOperacion.codoperaCont, 
    dbo.det_scrap.scd_cantidad AS Pzs, 
    dbo.DefectosXOperacion.codopera, dbo.operaciones.descripcion, 
    dbo.Defectos.descripcion AS Defecto, dbo.det_scrap.scd_item, 
    dbo.Cap_scrap.sc_folio, dbo.Cap_scrap.sc_tscrap, 
    dbo.Cap_scrap.sc_Etiqueta, dbo.Cap_scrap.sc_status, 
    dbo.Cap_scrap.reg_imp, dbo.Cap_scrap.impreso
FROM dbo.det_scrap INNER JOIN
    dbo.DefectosXOperacion ON 
    dbo.det_scrap.IdOperError = dbo.DefectosXOperacion.IdOperError INNER
     JOIN
    dbo.Cap_scrap ON 
    dbo.det_scrap.sc_folio = dbo.Cap_scrap.sc_folio INNER JOIN
    dbo.operaciones ON 
    dbo.DefectosXOperacion.codopera = dbo.operaciones.codopera INNER JOIN
    dbo.Defectos ON 
    dbo.det_scrap.coderror = dbo.Defectos.coderror
WHERE (dbo.Cap_scrap.sc_fecha = ?) AND (dbo.Cap_scrap.codlinea = ?) 
    AND (dbo.Cap_scrap.codturno = ?) AND (dbo.Cap_scrap.Oma_Id = ?) 
    AND (dbo.Cap_scrap.sc_status = N'P') AND 
    (dbo.Cap_scrap.sc_Etiqueta = ?) AND (dbo.Cap_scrap.sc_tscrap = ?) 
    AND (dbo.Cap_scrap.reg_imp = ?) AND (dbo.Cap_scrap.impreso = ?)
ORDER BY dbo.operaciones.descripcion, dbo.Defectos.descripcion