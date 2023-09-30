Option Explicit

Private aCnpj As String
Private aRazaoSocial As String
Private aAbertura As Date
Private aMatriz As Boolean



Public Property Let CNPJ(pCnpj As String): aCnpj = pCnpj: End Property
Public Property Get CNPJ() As String: CNPJ = aCnpj: End Property

Public Property Let RazaoSocial(pRazaoSocial As String): aRazaoSocial = pRazaoSocial: End Property
Public Property Get RazaoSocial() As String: RazaoSocial = aRazaoSocial: End Property

Public Property Let Abertura(pAbertura As Date): aAbertura = pAbertura: End Property
Public Property Get Abertura() As Date: Abertura = aAbertura: End Property

Public Property Let Matriz(pMatriz As Boolean): aMatriz = pMatriz: End Property
Public Property Get Matriz() As Boolean: Matriz = aMatriz: End Property



Public Sub LimparDados()
  aCnpj = ""
  aRazaoSocial = ""
  aAbertura = 0
  aMatriz = False
End Sub