Option Explicit

Private aCnpj As String
Private aRazaoSocial As String
Private aAbertura As Date
Private aMatriz As Boolean
Private aValido As Boolean
Private aEndereco As Endereco

Public Event CnpjValido(pCnpj As String, pValido As Boolean)



Public Property Let CNPJ(pCnpj As String)
  aValido = CnpjValidation(pCnpj)
  aCnpj = pCnpj
  Raise CnpjValidation(aCnpj, aValido)
End Property
Public Property Get CNPJ() As String: CNPJ = aCnpj: End Property

Public Property Let RazaoSocial(pRazaoSocial As String): aRazaoSocial = pRazaoSocial: End Property
Public Property Get RazaoSocial() As String: RazaoSocial = aRazaoSocial: End Property

Public Property Let Abertura(pAbertura As Date): aAbertura = pAbertura: End Property
Public Property Get Abertura() As Date: Abertura = aAbertura: End Property

Public Property Let Matriz(pMatriz As Boolean): aMatriz = pMatriz: End Property
Public Property Get Matriz() As Boolean: Matriz = aMatriz: End Property

Public Property Get Valido() As String: Valido = aValido: End Property

Public Property Get Endereco() As Endereco: Set Endereco = aEndereco: End Property



Public Sub LimparDados()
  aCnpj = ""
  aRazaoSocial = ""
  aAbertura = 0
  aMatriz = False
  aValido = False
  aEndereco.LimparDados
End Sub

Private Function CnpjValidation(ByVal pNrDoc As String) As Boolean
'---------------------------------------------------------------------------------------
' Procedure   : VerificarCNPJ (Nome Original) | CnpjValidation (Nome Modificado)
' Programador : Reinaldo Coral (Site: http://www.exceldoseujeito.com.br)
' Link        : http://www.exceldoseujeito.com.br/2008/12/18/validar-cpf-cnpj-e-titulo-de-eleitor-parte-i/
'---------------------------------------------------------------------------------------
  Dim d(14), DV(2), iDV, iDig, xFator As Integer
  ' Realiza um laço pelos dígitos do documento.
  For iDig = 1 To VBA.Len(pNrDoc)
    d(iDig) = VBA.CInt(VBA.Mid(pNrDoc, iDig, 1))
  Next iDig
  ' Aqui é executado o calculo para o Dígito Verificador.
  For iDV = 1 To 2
    ' Laço do Fator de Multiplicação.
    For iDig = 1 To VBA.Len(pNrDoc) - (3 - iDV)
      ' Define o valor da variável xFator de acordo
      ' com o Dígito Verificador e o Dígito Atual.
      If iDV = 1 Then xFator = VBA.IIf(iDig <= 4, 5, -3)
      If iDV = 2 Then xFator = VBA.IIf(iDig <= 5, 4, -4)
      ' Realiza o somatório acumulativo do cálculo.
      DV(iDV) = DV(iDV) + d(iDig) * (iDig + xFator)
    Next iDig
    ' Obtem o reto.
    DV(iDV) = VBA.IIf(DV(iDV) Mod 11 = 10, 0, DV(iDV) Mod 11)
  Next iDV
  ' Verifica se os valores dos Dígitos Calculados,
  ' coincidem com os Dígitos Informados no documento.
  CnpjValidation = VBA.IIf(VBA.CInt(DV(1)) = VBA.CInt(VBA.Mid(pNrDoc, 13, 1)) And _
                           VBA.CInt(DV(2)) = VBA.CInt(VBA.Mid(pNrDoc, 14, 1)), True, False)
End Function

Private Sub Class_Initialize(): Set aEndereco = Endereco: End Sub
Private Sub Class_Terminate(): Set aEndereco = Nothing: End Sub
