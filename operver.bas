Attribute VB_Name = "Module1"
Type oper_aciones
    CTA As String * 6
    descr As String * 30
    fe As String * 2
    impte As String * 16
    identi As String * 1
    Real As String * 9
End Type

Type ope
  CTA As String * 6
  reda As String * 30
  fecha As String * 2
  impo As String * 16
  clav As String * 1
  Real As String * 9
End Type
Type otr
  CTA As String * 6
  reda As String * 30
  fecha As String * 2
  impo As String * 16
  clav As String * 5
  Real As String * 5
End Type
Type Mayor
  B1 As String * 6
  B2 As String * 32
  B3 As String * 16
  B4 As String * 5
  B5 As String * 5
End Type
Type Auxiliar
  C1 As String * 6
  C2 As String * 32
  C3 As String * 16
  C4 As String * 5
  C5 As String * 5
End Type
Type RG
      DE As Integer
      A As Integer
      Ubic As Integer
End Type
Type DAT_OS
    D1 As String * 64
    D2 As String * 60
    D3 As String * 45
    No_arch As String * 15
    a_o As String * 5
    others1  As String * 25
    UltimaPol As String * 5
    UltimoReg As String * 5
    others As String * 12
End Type
Public Datos As DAT_OS
Public opera As oper_aciones
Public oper As ope, oper1 As ope, Real As ope
Public otros As otr, otros1 As otr
Public CATMAY As Mayor
Public CATAUX As Auxiliar
Public Real1 As otr
Public r As Long, I As Long
Public Orgs As RG, Dest As RG, CM As Long, DM As Long
Public Mm(15) As String * 20, dd(15) As Integer, m_m As Integer
Public dia As Integer
Public Mvtodebe As Currency, Mvtohaber As Currency, SALDO As Currency
