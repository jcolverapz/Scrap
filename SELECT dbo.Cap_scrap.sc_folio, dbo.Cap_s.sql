SELECT dbo.Cap_scrap.sc_folio, dbo.Cap_scrap.codturno, 
    dbo.Cap_scrap.codlinea, dbo.Cap_scrap.[of], 
    dbo.Cap_scrap.sc_status, dbo.Cap_scrap.sc_Etiqueta, 
    dbo.Cap_scrap.Oma_Id, dbo.det_scrap.scd_item, 
    dbo.det_scrap.coderror, dbo.det_scrap.proceso, 
    dbo.det_scrap.IdOperError, dbo.Cap_scrap.sc_tipo, 
    dbo.Cap_scrap.codvidrio, dbo.det_scrap.scd_obs, 
    dbo.Especificaciones.NoDeParte, dbo.Cap_scrap.sc_fecha, 
    dbo.Cap_scrap.reg_imp, dbo.Cap_scrap.impreso
FROM dbo.Cap_scrap INNER JOIN
    dbo.det_scrap ON 
    dbo.Cap_scrap.sc_folio = dbo.det_scrap.sc_folio INNER JOIN
    dbo.Especificaciones ON 
    dbo.Cap_scrap.[of] = dbo.Especificaciones.[OF]
WHERE 
(Cap_scrap.Oma_Id = ?) AND 
(Cap_scrap.codlinea = ?) AND 
(det_scrap.coderror = ?) AND 
(det_scrap.IdOperError = ?) AND 
(Especificaciones.NoDeParte = ?) AND 
(Cap_scrap.sc_Etiqueta = ?) AND
(Cap_scrap.codturno = ?) AND 
(Cap_scrap.sc_fecha = ?) AND 
(Cap_scrap.sc_status = ?) AND 
(Cap_scrap.reg_imp = ?) AND 
(Cap_scrap.impreso = ?)

WHERE (Cap_scrap.Oma_Id = ?) AND (Cap_scrap.codlinea = ?) AND 
    (det_scrap.coderror = ?) AND (det_scrap.IdOperError = ?) AND 
    (Cap_scrap.sc_Etiqueta = ?) AND (Cap_scrap.codturno = ?) AND 
    (Cap_scrap.sc_fecha = ?) AND (Cap_scrap.sc_status = ?) AND 
    (Cap_scrap.reg_imp = ?) AND (Cap_scrap.impreso = ?)