VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12510
   LinkTopic       =   "Form2"
   ScaleHeight     =   7815
   ScaleWidth      =   12510
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtdireccion 
      Height          =   495
      Left            =   4560
      TabIndex        =   10
      Top             =   4920
      Width           =   2535
   End
   Begin VB.TextBox txtcorreo 
      Height          =   495
      Left            =   4560
      TabIndex        =   9
      Top             =   4200
      Width           =   2535
   End
   Begin VB.TextBox txtnombre 
      Height          =   495
      Left            =   4560
      TabIndex        =   8
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox txtcedula 
      Height          =   495
      Left            =   4560
      TabIndex        =   7
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Despues"
      Height          =   495
      Left            =   9720
      TabIndex        =   6
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Antes"
      Height          =   495
      Left            =   7560
      TabIndex        =   5
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   7560
      TabIndex        =   4
      Top             =   5280
      Width           =   4095
   End
   Begin VB.CommandButton cmdborrar 
      Caption         =   "Borrar"
      Height          =   495
      Left            =   7560
      TabIndex        =   3
      Top             =   4440
      Width           =   4095
   End
   Begin VB.CommandButton cmdguardar 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   7560
      TabIndex        =   2
      Top             =   3600
      Width           =   4095
   End
   Begin VB.CommandButton cmdnuevo 
      Caption         =   "Nuevo"
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   2760
      Width           =   4095
   End
   Begin VB.Label Label5 
      Caption         =   "Dirección "
      Height          =   495
      Left            =   1680
      TabIndex        =   14
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Correo"
      Height          =   495
      Left            =   1800
      TabIndex        =   13
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Cédula"
      Height          =   495
      Left            =   1800
      TabIndex        =   12
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre"
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Registro de Datos"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      TabIndex        =   0
      Top             =   480
      Width           =   6735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
