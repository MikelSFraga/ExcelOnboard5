Option Explicit

Private aLogradouro As String
Private aNumero As String
Private aComplemento As String
Private aBairro As String
Private aCidade As String
Private aUf As String
Private aCep As String




Public Property Let Logradouro(pLogradouro As String): aLogradouro = pLogradouro: End Property
Public Property Get Logradouro() As String: Logradouro = aLogradouro: End Property

Public Property Let Numero(pNumero As String): aNumero = pNumero: End Property
Public Property Get Numero() As String: Numero = aNumero: End Property

Public Property Let Complemento(pComplemento As String): aComplemento = pComplemento: End Property
Public Property Get Complemento() As String: Complemento = aComplemento: End Property

Public Property Let Bairro(pBairro As String): aBairro = pBairro: End Property
Public Property Get Bairro() As String: Bairro = aBairro: End Property

Public Property Let Cidade(pCidade As String): aCidade = pCidade: End Property
Public Property Get Cidade() As String: Cidade = aCidade: End Property

Public Property Let Uf(pUf As String): aUf = pUf: End Property
Public Property Get Uf() As String: Uf = aUf: End Property

Public Property Let Cep(pCep As String): aCep = pCep: End Property
Public Property Get Cep() As String: Cep = aCep: End Property

Public Property Get Completo() As String
  Completo = "R./Av." & aLogradouro & _
              ", " & IIf(aNumero = "", "S/N", aNumero) & _
              IIf(aComplemento = "", "", " - " & aComplemento) & _
              IIf(aBairro = "", " ", " - " & aBairro) & _
              IIf(aCidade = "", " ", " - " & aCidade) & _
              IIf(aUf = "", "", "/" & aUf) & _
              IIf(aCep = "", "", " Cep: " & aCep)
End Property



Public Sub LimparDados()
  aLogradouro = ""
  aNumero = ""
  aComplemento = ""
  aBairro = ""
  aCidade = ""
  aUf = ""
  aCep = ""
End Sub
