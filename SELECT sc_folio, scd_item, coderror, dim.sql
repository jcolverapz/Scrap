SELECT sc_folio, scd_item, coderror, dimension_item, dimension_x, 
    dimension_y, status, fecha_hora, etiqueta
FROM dbo.Tbl_Dimension_Scrap
WHERE (sc_folio = ?)