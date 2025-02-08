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
   Begin VB.Frame frmVariable 
      Caption         =   "Variable <"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3105
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   7935
      Begin VB.ListBox lstVariable 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2505
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   7695
      End
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
      Height          =   465
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   7935
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
      TabIndex        =   5
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame frmTerminal 
      Caption         =   "Terminal <"
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
         Left            =   3840
         TabIndex        =   4
         Top             =   2640
         Width           =   1935
      End
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
         TabIndex        =   3
         Top             =   2640
         Width           =   3615
      End
      Begin VB.CommandButton cmdClearTerminal 
         Caption         =   "Limpar"
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
         Left            =   5880
         TabIndex        =   2
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox txtTerminal 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2175
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   360
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
Dim buffer As String
Dim arrayName(9) As String

Private Sub Form_Load()
    
    Titulo = App.Title & "   " & "v" & App.Major & "." & App.Minor & "." & App.Revision
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
    
    ' Informação de funções a serem programadas no arduino
    lstVariable.ToolTipText = "Selecione e dê um duplo click para criar um nome para a variável."
    frmVariable.ToolTipText = "Exemplo para arduino -> Serial.println(""V0:"" + String(value_variable); delay(100); " & _
                              "é interessante o uso do delay em seguida."
    
    ' ' Ajustes inciais
    cmdConnectPort.Enabled = True
    frmTerminal.Visible = False
    frmVariable.Visible = True
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
        frmTerminal.Enabled = False
        frmVariable.Enabled = False
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
            buffer = buffer + Data
            'Debug.Print buffer
            
            ' Atualiza valores de Output e Analog
            If Mid(buffer, 3, 1) = ":" Then
                If Mid(buffer, 1, 1) = "V" Then
                    Call updateVariable(buffer) ' Variable
                End If
            End If
            
            ' Limpa o buffer se receber uma nova linha
            If Right(buffer, 1) = vbLf Then
                Debug.Print buffer
                buffer = Empty
            End If
            
            ' Atualiza terminal serial
            With txtTerminal
                .SelStart = Len(txtTerminal.Text)
                .SelText = Data
                .SelStart = Len(txtTerminal.Text)
            End With
            
    End Select
    
Exit Sub

Erro:
    If Err = 13 Then Exit Sub
    MsgBox "Erro " & Err & ". " & Error, vbCritical, "DALCOQUIO AUTOMAÇÃO"

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

Private Sub lstVariable_DblClick()
    ' Obtém o item selecionado
    Dim selectedItem As String
    selectedItem = lstVariable.List(lstVariable.ListIndex)
    If selectedItem = Empty Then Exit Sub
    
    Dim resposta As String
    resposta = InputBox("Digite um nome para esta variável.", "DALÇÓQUIO AUTOMAÇÃO", arrayName(lstVariable.ListIndex))
    arrayName(lstVariable.ListIndex) = resposta
    
    lstVariable.List(lstVariable.ListIndex) = selectedItem & arrayName(lstVariable.ListIndex)
    
    Call ClearBufferSerial

End Sub

Private Sub updateVariable(Data As String)
    Dim searchPrefix As String
    Dim i As Integer
    Dim foundIndex As Integer

    ' Extrai a parte antes do ":" para usar como prefixo de busca
    searchPrefix = Left(Data, InStr(Data, ":"))

    ' Define um índice inicial para indicar que nenhum item foi encontrado
    foundIndex = -1

    ' Procura no ListBox pelo item que começa com o prefixo
    For i = 0 To lstVariable.ListCount - 1
        If Left(lstVariable.List(i), Len(searchPrefix)) = searchPrefix Then
            ' Encontrou o item correspondente
            foundIndex = i
            Exit For
        End If
    Next i

    ' Verifica se o item foi encontrado
    If foundIndex <> -1 Then
        ' Atualiza o item no ListBox com o valor recebido
        lstVariable.List(foundIndex) = Data & " " & Time & " " & arrayName(foundIndex)
    End If

End Sub

Private Sub cmdReset_Click()
    ' Reset hardware via software
    MSComm1.DTREnable = False
    MSComm1.DTREnable = True
    
    ' Limpa listbox e terminal
    lstVariable.Clear
    txtTerminal.Text = Empty
    
    ' Lista Variable
    For i = 0 To 9
        lstVariable.AddItem "V" & i & ":"
    Next i
    
End Sub

Private Sub clearValue()
    shpInput.BackColor = vbBlack

End Sub

Private Sub cmdClearTerminal_Click()
    txtTerminal.Text = Empty
    
End Sub

Private Sub mTerminal_Click()
    If frmVariable.Visible = True Then
        frmVariable.Visible = False
        frmTerminal.Visible = True
        mTerminal.Caption = "Variable"
    Else
        frmTerminal.Visible = False
        frmVariable.Visible = True
        mTerminal.Caption = "Terminal"
    End If
    
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

