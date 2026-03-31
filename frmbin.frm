VERSION 5.00
Begin VB.Form frmbin 
   BackColor       =   &H00C0C0C0&
   Caption         =   "ascbin&binasc"
   ClientHeight    =   2235
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   ScaleHeight     =   2235
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "decode"
      Height          =   495
      Left            =   1890
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1365
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "code"
      Height          =   495
      Left            =   195
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1380
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1125
      TabIndex        =   0
      Top             =   555
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   735
      TabIndex        =   3
      Top             =   45
      Width           =   1995
   End
End
Attribute VB_Name = "frmbin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Click()
Dim resto As Single, valor As Single, bininvert As String, quantnum As Integer, numbin As String, n As Integer
'resto calcula o resto da divisão
'valor é o número a ser trabalhado
'bininvert número binário invertido
'quantnum quantidade de números
'numbin número binário
'n uutilizado em loop
On Error GoTo fim
resto = Asc(txt.Text)
valor = Asc(txt.Text)
bininvert = ""
Do
    resto = (resto Mod 2)
    valor = (valor \ 2)
    bininvert = bininvert & resto
    resto = valor
    valor = valor
Loop While valor > 1
bininvert = bininvert & valor
quantnum = Len(bininvert)
For n = 0 To (quantnum - 1)
    numbin = numbin & Mid(bininvert, (quantnum - n), 1)
Next
lbl.Caption = numbin
fim:
End Sub

Private Sub cmdd_Click()
Dim numasc, numbin As String, quantnum As Integer, bimmult As Integer, n As Variant
'numasc valor asc
'numbin valor binário
'quantnum quantidade de números
'bimmult binario multiplicado
'n utilizado em loop
numbin = lbl.Caption
quantnum = Len(numbin)
For n = 0 To (quantnum - 1)
    On Error GoTo fim
    bimmult = Mid(numbin, (quantnum - n), 1)
    If bimmult = 1 Then
        numasc = numasc + (bimmult * (2 ^ n))
    End If
Next
lbl.Caption = Chr(numasc)
fim:
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)

txt.Text = ""

End Sub
