VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frm_server 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CHAT Servidor"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7545
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleMode       =   0  'User
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_reiniciar 
      Caption         =   "Reiniciar"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmd_apagar 
      Caption         =   "Apagar"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox txt_log 
      Height          =   3975
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7011
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frm_server.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txt_name 
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   6120
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_send 
      Caption         =   "Enviar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox txt_mensaje 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   5160
      Width           =   5655
   End
   Begin VB.CommandButton cmd_iniciar 
      Caption         =   "Iniciar"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbl_estado 
      Caption         =   "Label2"
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "frm_server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_apagar_Click()
    Winsock.Close
    txt_log.Text = txt_log.Text & vbCrLf & "El cliente se desconecto" & vbCrLf
    If (txt_log.Text <> "") Then
        txt_log.Text = ""
        txt_log.Text = "Servidor apagado"
    Else
        
    End If
    lbl_estado.Caption = Winsock.State
End Sub

Private Sub cmd_reiniciar_Click()
    Winsock.Close
    Winsock.LocalPort = 60000
    Winsock.Listen
    If (txt_log.Text <> "") Then
        txt_log.Text = ""
        txt_log.Text = "Servidor reiniciado" & vbCrLf
    Else
        
    End If
    lbl_estado.Caption = Winsock.State
End Sub

Private Sub Form_Load()
    lbl_estado.Caption = Winsock.State
End Sub

Private Sub txt_name_LostFocus()
    If (txt_name.Text <> "") Then
        txt_name.BackColor = vbWhite
    End If
End Sub

Private Sub cmd_iniciar_Click()
    Winsock.Close
    Winsock.LocalPort = 60000
    Winsock.Listen
    If (txt_log.SelText = "") Then
        txt_log.SelText = "Servidor inciado" & vbCrLf
    Else
        
    End If
    lbl_estado.Caption = Winsock.State
End Sub

Private Sub cmd_send_Click()
    If (txt_name.Text = "") Then
        txt_name.BackColor = vbRed
        txt_name.SetFocus
        txt_mensaje.Text = ""
        MsgBox "Ingresa un nombre primero", vbInformation, ""
    Else
        If (txt_mensaje.Text = "") Then
            MsgBox "No puedes enviar mensajes vacios", vbExclamation
            txt_mensaje.SetFocus
        Else
                txt_log.SelColor = vbBlue
                Winsock.SendData txt_name & "(" & Time & "): " & txt_mensaje.Text
                txt_log.SelText = txt_log.SelText & vbCrLf & txt_name & "(" & Time & "): " & txt_mensaje.Text
                txt_mensaje.Text = ""
                txt_mensaje.SetFocus
                txt_log.SelStart = Len(txt_log)
        End If
    End If
End Sub

Private Sub Winsock_ConnectionRequest(ByVal requestID As Long)
    Winsock.Close
    Winsock.Accept requestID
    txt_log.SelText = "Cliente conectado. IP : " & Winsock.RemoteHostIP & vbCrLf
    lbl_estado.Caption = Winsock.State
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    Dim msjrecibido As String
    Winsock.GetData msjrecibido, vbString
    txt_log.SelColor = vbRed
    txt_log.SelText = txt_log.SelText & vbCrLf & msjrecibido
    txt_log.SelStart = Len(txt_log)
End Sub

Private Sub txt_mensaje_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        cmd_send_Click
    End If
End Sub

Private Sub winsock_Close()
    Winsock.Close
    txt_log.Text = txt_log.Text & vbCrLf & "El cliente se desconecto" & vbCrLf
    txt_log.SelText = ""
    cmd_iniciar_Click
End Sub


