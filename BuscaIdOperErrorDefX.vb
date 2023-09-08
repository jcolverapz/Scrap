CNN.CmdBuscaIdOperErrorDefX (TipoScrap), (CodScrap)
     If CNN.rsCmdBuscaIdOperErrorDefX.EOF <> True Then
    IdOperError = CNN.rsCmdBuscaIdOperErrorDefX!IdOperError
Else
    IdOperError = 0
End If
CNN.rsCmdBuscaIdOperErrorDefX.Close

'modificacion
SELECT DefectosXOperacion.IdOperError, DefectosXOperacion.codopera, 
    DefectosXOperacion.coderror, tipodef.codtipoe
FROM DefectosXOperacion INNER JOIN
    Defectos ON 
    DefectosXOperacion.coderror = Defectos.coderror INNER JOIN
    tipodef ON Defectos.codtipoe = tipodef.codtipoe
WHERE (DefectosXOperacion.coderror = ?) AND 
    (DefectosXOperacion.codopera = ?)