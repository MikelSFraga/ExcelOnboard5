Option Explicit

Private aCnpj As String

Public Property Let CNPJ(pCnpj As String): aCnpj = pCnpj: End Property
Public Property Get CNPJ() As String: CNPJ = aCnpj: End Property