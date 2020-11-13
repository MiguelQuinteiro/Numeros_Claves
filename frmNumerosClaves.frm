VERSION 5.00
Begin VB.Form frmNumerosClaves 
   AutoRedraw      =   -1  'True
   Caption         =   "Números Claves"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNumero 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   8280
      TabIndex        =   2
      Text            =   "1000"
      Top             =   720
      Width           =   1215
   End
   Begin VB.ListBox lstNumerosClaves 
      Height          =   6690
      Left            =   8280
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "Calcular"
      Height          =   495
      Left            =   8280
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmNumerosClaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim miN As Long

Private Sub cmdCalcular_Click()
  Dim i As Long
  miN = Val(txtNumero.Text)
  lstNumerosClaves.Clear

  Open "Claves.txt" For Output As 1

  For i = 1 To miN
    If CalculaClave(i) Then
      lstNumerosClaves.AddItem i
      'Print #1, i
    End If
  Next i

  Close #1
End Sub

Public Function CalculaClave(ByVal pN As Long) As Boolean
  Dim i As Long
  Dim miCantidadFactores As Integer
  Dim miElemento(4) As Long
  CalculaClave = False
  miCantidadFactores = 0
  For i = 1 To pN
    If pN / i = Int(pN / i) Then
      miCantidadFactores = miCantidadFactores + 1
      If miCantidadFactores <= 4 Then
        miElemento(miCantidadFactores) = i
      Else
        i = pN
      End If
    End If
  Next i
  If miCantidadFactores = 4 Then

    ' Condicion principal
    If Primo(miElemento(2)) And Primo(miElemento(3)) And miElemento(2) <> 2 And miElemento(2) <> 3 And miElemento(2) <> 5 Then
      Print Tab(20), pN, miElemento(2), miElemento(3)
      CalculaClave = True

      Print #1, (miElemento(2) * miElemento(3)), miElemento(2), miElemento(3)
      'Print #1, i


    End If

  End If
End Function


' Funcion para saber si un numero es primo
Public Function Primo(ByVal pN As Integer) As Boolean
  Dim i As Integer
  Primo = True
  If pN = 1 Then
    Primo = False
  Else
    For i = 2 To Sqr(pN)
      If (pN / i) = Int(pN / i) Then
        Primo = False
      End If
    Next i
  End If
End Function


