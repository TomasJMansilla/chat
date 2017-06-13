VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Super Mega Chat"
   ClientHeight    =   5910
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   10290
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu cliente 
      Caption         =   "Cliente"
   End
   Begin VB.Menu servidor 
      Caption         =   "Servidor"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cliente_Click()
    frm_cliente.Show
End Sub

Private Sub servidor_Click()
    frm_server.Show
End Sub
