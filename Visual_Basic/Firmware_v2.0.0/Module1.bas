Attribute VB_Name = "Module1"
Type DCB
  DCBlength As Long
  BaudRate As Long
  fBitFields As Long
  wReserved As Integer
  XonLim As Integer
  XoffLim As Integer
  ByteSize As Byte
  Parity As Byte
  StopBits As Byte
  XonChar As Byte
  XoffChar As Byte
  ErrorChar As Byte
  EofChar As Byte
  EvtChar As Byte
  wReserved1 As Integer
End Type

Type COMMCONFIG
  dwSize As Long
  wVersion As Integer
  wReserved As Integer
  dcbx As DCB
  dwProviderSubType As Long
  dwProviderOffset As Long
  dwProviderSize As Long
  wcProviderData As Byte
End Type

Declare Function GetDefaultCommConfig Lib "kernel32" _
Alias "GetDefaultCommConfigA" (ByVal lpszName As String, _
lpCC As COMMCONFIG, lpdwSize As Long) As Long

Public Function DetectaPortaCOM(port As Integer) As Long
'retorna zero se a porta com não existir
Dim cc As COMMCONFIG, ccsize As Long

ccsize = LenB(cc)

DetectaPortaCOM = GetDefaultCommConfig("COM" + Trim(Str(port)) + Chr(0), cc, ccsize)

End Function



