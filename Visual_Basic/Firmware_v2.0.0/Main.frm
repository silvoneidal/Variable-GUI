VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4830
   ClientLeft      =   1350
   ClientTop       =   2970
   ClientWidth     =   8190
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboSendData 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   3015
   End
   Begin VB.CommandButton cmdSendData 
      Caption         =   "Enviar"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3360
      TabIndex        =   7
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5760
      TabIndex        =   6
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   840
      Top             =   4440
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame frmConnect 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   7935
      Begin VB.CommandButton cmdScanPort 
         Caption         =   "Scanear"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3120
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmdConnectPort 
         Caption         =   "Conectar"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   5520
         TabIndex        =   4
         Top             =   240
         Width           =   2295
      End
      Begin VB.ComboBox cboBaudRate 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cboCommPort 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame frmTerminal 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   7935
      Begin VB.TextBox txtTerminal 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "Main.frx":6852
         Top             =   240
         Width           =   7695
      End
   End
   Begin VB.Shape shpConnect 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   -120
      Top             =   4680
      Width           =   8415
   End
   Begin VB.Menu mTerminal 
      Caption         =   "Terminal"
   End
   Begin VB.Menu mGerenciador 
      Caption         =   "Gerenciador de Dispositivos"
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Microsoft Comm Control 6.0
' Microsoft Windows Common Controls 6.0 (SP6)

' Para uso de sleep
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Dim Titulo As String
Dim index As Integer
Dim Buffer As String
Dim Debugger As Boolean
Dim arrayVariable(9) As String
Dim arrayName(9) As String

Private Sub Form_Load()
    
    Titulo = App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision & "   by   DALÇÓQUIO AUTOMAÇÃO"
    Me.Caption = Titulo

On Error GoTo Erro
        
    ' Adiciona lista de baudrate
    cboBaudRate.AddItem "1200"
    cboBaudRate.AddItem "2400"
    cboBaudRate.AddItem "4800"
    cboBaudRate.AddItem "9600"
    cboBaudRate.AddItem "19200"
    cboBaudRate.AddItem "38400"
    cboBaudRate.AddItem "57600"
    cboBaudRate.AddItem "115200"
    cboBaudRate.ListIndex = 3
    
    ' Busca portas disponiveis
    Call cmdScanPort_Click
    
    ' Ajustes inciais
    cmdConnectPort.Enabled = True
    shpConnect.BackColor = vbRed
    
    ' Inicia com reset de valores
    Call cmdReset_Click
    
Exit Sub

Erro:
    MsgBox "Erro " & Err & ". " & Error, vbCritical, "DALCOQUIO AUTOMAÇÃO"
    
End Sub

Private Sub cmdScanPort_Click()
    Dim i As Integer
    
On Error GoTo Erro

    'Detecta portas disponiveis
    cmdScanPort.Caption = "Scan..."
    cboCommPort.Clear
    For i = 1 To 32
        If DetectaPortaCOM(i) <> 0 Then
            cboCommPort.AddItem "COM" & i
        End If
    Next
    cboCommPort.ListIndex = 0
    cmdScanPort.Caption = "Scanear"
        
Exit Sub
    
Erro:
    MsgBox "Erro " & Err & ". " & Error, vbCritical, "DALCOQUIO AUTOMAÇÃO"

End Sub

Private Sub cmdConnectPort_Click()
On Error GoTo Erro

    ' Conectar
    If cmdConnectPort.Caption = "Conectar" Then
        cmdConnectPort.Caption = "Desconectar"
        MSComm1.CommPort = Mid(cboCommPort.Text, 4, 2)
        MSComm1.Settings = cboBaudRate.Text & "n,8,1"
        MSComm1.RThreshold = 1
        MSComm1.DTREnable = True
        MSComm1.RTSEnable = True
        MSComm1.PortOpen = True
        cboCommPort.Enabled = False
        cboBaudRate.Enabled = False
        cmdScanPort.Enabled = False
        shpConnect.BackColor = vbGreen
        Me.Caption = "Conectado na COM" & MSComm1.CommPort & "," & MSComm1.Settings
    ' Desconectar
    Else
        cmdConnectPort.Caption = "Conectar"
        MSComm1.PortOpen = False
        cboCommPort.Enabled = True
        cboBaudRate.Enabled = True
        cmdScanPort.Enabled = True
        shpConnect.BackColor = vbRed
        Call cmdReset_Click
        Me.Caption = Titulo
    End If
    
    
Exit Sub

Erro:
    ' Erro relacionados a porta serial
    If Err = 8005 Or Err = 8002 Or Err = 8020 Then
        cmdConnectPort.Caption = "Conectar"
        cboCommPort.Enabled = True
        cboBaudRate.Enabled = True
        cmdScanPort.Enabled = True
        shpConnect.BackColor = vbRed
        Me.Caption = Titulo
    End If
    
    MsgBox "Erro " & Err & ". " & Error, vbCritical, "DALCOQUIO AUTOMAÇÃO"
    
End Sub

Private Sub MSComm1_OnComm()
    
On Error GoTo Erro

    If MSComm1.PortOpen = False Then Exit Sub
    If pause = True Then Exit Sub

    Select Case MSComm1.CommEvent
        Case comEvReceive
            ' Recebe os dados da serial
            Dim Data As String
            Data = MSComm1.Input
            Buffer = Buffer + Data
            
            If Debugger = True Then
                ' Atualiza debug de variáveis
                Dim i As Integer
                For i = 0 To 9
                    If Mid(Buffer, 1, 3) = "V" & i & ":" Then
                        Call updateVariable(i, Buffer) ' Variable
                        Exit For
                    End If
                Next i
            Else
                ' Atualiza terminal serial
                With txtTerminal
                    .SelStart = Len(txtTerminal.Text)
                    .SelText = Data
                    .SelStart = Len(txtTerminal.Text)
                End With
            End If
            
            ' Limpa o buffer se receber uma nova linha
            If Right(Buffer, 1) = vbLf Then
                Buffer = Empty
            End If
            
    End Select
    
Exit Sub

Erro:
    If Err = 380 Then Exit Sub
    MsgBox "Erro " & Err & ". " & Error, vbCritical, "DALCOQUIO AUTOMAÇÃO"

End Sub

Private Sub updateVariable(index As Integer, Buffer As String)
    ' Remove retorno de carro do buffer
    Buffer = Replace(Buffer, vbCr, "")
    
    ' Atualiza valor da variável no array
    Dim X As Integer
    For X = 0 To 9
        If X = index Then
            arrayVariable(X) = Buffer & " " & Time & " " & arrayName(X)
            Exit For
        End If
    Next X
    
    ' Limpa o terminal antes de atualizar
    txtTerminal.Text = Empty
    
    ' Atualiza todas as variáveis no terminal
    Dim Y As Integer
    For Y = 0 To 9
            txtTerminal.Text = txtTerminal.Text + arrayVariable(Y) & vbCrLf
    Next Y

End Sub

Private Sub ClearBufferSerial()
  Dim Data As String

  ' Verifica se há dados disponíveis
  Do While MSComm1.InBufferCount > 0
    ' Lê todos os dados disponíveis
    Data = MSComm1.Input
    ' Descarta os dados (opcional: você poderia processá-los aqui se necessário)
  Loop
End Sub

Private Sub cmdSendData_Click()
    If MSComm1.PortOpen = False Then Exit Sub
    If cboSendData.Text = "" Then Exit Sub
    
    ' Envia dado pela serial
    MSComm1.Output = cboSendData.Text
    
    'atualiza cboSend
    cboSendData.AddItem cboSendData.Text
    deleteDuplicados
    cboSendData.Text = Empty
    
End Sub

Private Sub deleteDuplicados()
    Dim i As Integer, j As Integer
    
    For i = 0 To cboSendData.ListCount
        For j = i + 1 To cboSendData.ListCount
            If cboSendData.List(i) = cboSendData.List(j) Then
                cboSendData.RemoveItem (j)
                j = j - 1
            End If
        Next
    Next
    
End Sub

Private Sub cmdReset_Click()
    ' Reset hardware via software
    MSComm1.DTREnable = False
    MSComm1.DTREnable = True
    
    ' Limpa o buffer da serial
    Call ClearBufferSerial
    
    ' Limpa terminal da serial
    txtTerminal.Text = Empty
    
    ' Reset para terminal ou debug
    If Debugger = True Then
        For i = 0 To 9
            arrayVariable(i) = "V" & i & ":"
            txtTerminal.Text = txtTerminal.Text + arrayVariable(i) & vbCrLf
        Next i
    End If
    
End Sub

Private Sub clearValue()
    shpInput.BackColor = vbBlack

End Sub

Private Sub cmdClearTerminal_Click()
    txtTerminal.Text = Empty
    
End Sub

Private Sub mTerminal_Click()
    If mTerminal.Caption = "Debug" Then
        txtTerminal.Text = Empty
        mTerminal.Caption = "Terminal"
        txtTerminal.ToolTipText = Empty
        Debugger = False
    Else
        txtTerminal.Text = Empty
        mTerminal.Caption = "Debug"
        Debugger = True
        Call cmdReset_Click
        txtTerminal.ToolTipText = "Exemplo para arduino -> Serial.println(""V0:"" + String(value_variable); delay(100); " & _
                              "é recomendado uso do delay em seguida. Dê um duplo click para nomear a variável."
    End If
    
End Sub

Private Sub txtTerminal_DblClick()
    If mTerminal.Caption = "Terminal" Then Exit Sub
    
    ' Aguarda retorno do usário
    Dim retorno As String
    retorno = InputBox("Digite o index de sua variável, seguido do nome desejado. " & _
                       "Ex: V0:Nome_da_variável.", "DALÇÓQUIO AUTOMAÇÃO")
                  
    ' Verifica se variável é válida
    Dim teste As Boolean
    teste = False
    For i = 0 To 9
        If Left(retorno, 3) = "V" & i & ":" Then
            teste = True
            Exit For
        End If
    Next i
    If teste = False Then GoTo None
                       
    ' Se válida, atualiza lista de nomes
    If retorno <> Empty Then
        Dim index As Integer
        index = Mid(retorno, 2, 1)
        arrayName(index) = Mid(retorno, 4, Len(retorno))
    End If
    
None:
    ' Limpa o buffer da serial
    Call ClearBufferSerial

End Sub

Private Sub mGerenciador_Click()
    Shell ("cmd.exe /c devmgmt.msc")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Erro
    
    If MSComm1.PortOpen Then MSComm1.PortOpen = False
    
Exit Sub

Erro:
    MsgBox "Erro " & Err & ". " & Error, vbCritical, "DALCOQUIO AUTOMAÇÃO"
    
End Sub

'------------------------------------------------------------------------------------------
'Mid: Retorna o número especificado de caracteres de uma string.
'exemplo: mid(text1.text,1,5) -> retorna as letras 1,2,3,4,5 do text1.
'exemplo: mid(text1.text,20,5) -> retorna  as ultimas 5 letras iniciando da posicai 20 do text1.

'Left:Retorna o número especificado de caracteres a partir do início de uma string.
'exemplo: left(text1.text,3) -> retorna as 3 primeiras letras do text1.

'right:Retorna o número especificado de caracteres a partir do lado direito de uma string.
'exemplo: right(text1.text, 4) -> retorna as quatro últimas letras do text1.


' Função para verificar tempo de processo
'------------------------------------------------------------------------------------------
' Start tempo de processo
'Dim startTime As Double
'Dim endTime As Double
'Dim elapsedTime As Double
'startTime = Timer

' Loop de processo aqui...

' End tempo de processo
'endTime = Timer
'elapsedTime = endTime - startTime
'If elapsedTime > 2 Then txtData = Empty ' limpa txtData, pois houve algum erro.

