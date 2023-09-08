SELECT sc_folio, reg_imp, impreso, Fecha_imp, Maquina_registra, 
    codturno
FROM dbo.Cap_scrap
WHERE (sc_folio = ?)

SELECT Cap_scrap.sc_fecha, Cap_scrap.codlinea, Cap_scrap.codturno, 
    Cap_scrap.Oma_Id, SUM(det_scrap.scd_cantidad) AS Pzs, 
    det_scrap.coderror AS codError, Defectos.descripcion AS defecto, 
    det_scrap.sc_folio AS folio, det_scrap.scd_item, Cap_scrap.sc_folio, 
    Cap_scrap.sc_tscrap, Cap_scrap.sc_Etiqueta, Cap_scrap.impreso, 
    operaciones.descripcion, Cap_scrap.reg_imp
FROM det_scrap INNER JOIN
    Cap_scrap ON det_scrap.sc_folio = Cap_scrap.sc_folio INNER JOIN
    Defectos ON det_scrap.coderror = Defectos.coderror INNER JOIN
    operaciones ON det_scrap.proceso = operaciones.codopera
GROUP BY Cap_scrap.sc_fecha, Cap_scrap.codlinea, Cap_scrap.codturno, 
    Cap_scrap.Oma_Id, det_scrap.coderror, Defectos.descripcion, 
    det_scrap.sc_folio, det_scrap.scd_item, Cap_scrap.sc_folio, 
    Cap_scrap.sc_tscrap, Cap_scrap.sc_Etiqueta, Cap_scrap.impreso, 
    operaciones.descripcion, Cap_scrap.reg_imp
HAVING (Cap_scrap.sc_fecha = ?) AND (Cap_scrap.codlinea = ?) AND 
    (Cap_scrap.codturno = ?) AND (Cap_scrap.Oma_Id = ?) AND 
    (Cap_scrap.sc_Etiqueta = ?) AND (Cap_scrap.sc_tscrap = ?) AND 
    (Cap_scrap.reg_imp = ?) AND (Cap_scrap.impreso = ?)