VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
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
   Begin VB.ComboBox cmb_fondo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frm_server.frx":0000
      Left            =   120
      List            =   "frm_server.frx":0019
      TabIndex        =   7
      Text            =   "Fondo"
      Top             =   4680
      Width           =   1575
   End
   Begin VB.ComboBox cmb_letra 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frm_server.frx":0050
      Left            =   1800
      List            =   "frm_server.frx":0069
      TabIndex        =   6
      Text            =   "Color Letra"
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox txt_name 
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   5400
      Top             =   5400
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
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   5160
      Width           =   5655
   End
   Begin VB.CommandButton cmd_iniciar 
      Caption         =   "Iniciar"
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txt_log 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3615
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   960
      Width           =   7215
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
      TabIndex        =   4
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "frm_server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb_fondo_Click()
    If (cmb_fondo.Text = "Amarillo") Then
        txt_log.BackColor = vbYellow
    ElseIf (cmb_fondo.Text = "Azul") Then
        txt_log.BackColor = vbBlue
    ElseIf (cmb_fondo.Text = "Celeste") Then
        txt_log.BackColor = &HFFFF00
    ElseIf (cmb_fondo.Text = "Rojo") Then
        txt_log.BackColor = vbRed
    ElseIf (cmb_fondo.Text = "Negro") Then
        txt_log.BackColor = vbBlack
    ElseIf (cmb_fondo.Text = "Rosa") Then
        txt_log.BackColor = &HFF80FF
    ElseIf (cmb_fondo.Text = "Verde") Then
        txt_log.BackColor = vbGreen
    End If
End Sub

Private Sub txt_name_LostFocus()
    If (txt_name.Text <> "") Then
        txt_name.BackColor = vbWhite
    End If
End Sub

Private Sub cmb_letra_Click()
     If (cmb_letra.Text = "Amarillo") Then
        txt_log.ForeColor = vbYellow
    ElseIf (cmb_letra.Text = "Azul") Then
        txt_log.ForeColor = vbBlue
    ElseIf (cmb_letra.Text = "Celeste") Then
        txt_log.ForeColor = &HFFFF00
    ElseIf (cmb_letra.Text = "Rojo") Then
        txt_log.ForeColor = vbRed
    ElseIf (cmb_letra.Text = "Negro") Then
        txt_log.ForeColor = vbBlack
    ElseIf (cmb_letra.Text = "Rosa") Then
        txt_log.ForeColor = &HFF80FF
    ElseIf (cmb_letra.Text = "Verde") Then
        txt_log.ForeColor = vbGreen
    End If
End Sub

Private Sub cmd_iniciar_Click()
    Winsock.Close
    Winsock.LocalPort = 60000
    Winsock.Listen
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
            Winsock.SendData txt_name & "(" & Time & "): " & txt_mensaje.Text
            txt_log.Text = txt_log.Text & vbCrLf & txt_name & "(" & Time & "): " & txt_mensaje.Text
            txt_mensaje.Text = ""
            txt_mensaje.SetFocus
            txt_log.SelStart = Len(txt_log)
        End If
    End If
End Sub

Private Sub Winsock_ConnectionRequest(ByVal requestID As Long)
    Winsock.Close
    Winsock.Accept requestID
    txt_log.Text = "Cliente conectado. IP : " & Winsock.RemoteHostIP & vbCrLf
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    Dim msjrecibido As String
    Winsock.GetData msjrecibido, vbString
    txt_log.Text = txt_log.Text & vbCrLf & msjrecibido
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
End Sub

