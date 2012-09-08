VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Principal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INTERFAZ LAVADORA"
   ClientHeight    =   1905
   ClientLeft      =   1815
   ClientTop       =   2190
   ClientWidth     =   3225
   Icon            =   "LAVADORA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1905
   ScaleMode       =   0  'User
   ScaleWidth      =   3323.671
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2040
      TabIndex        =   13
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
      Max             =   4
   End
   Begin VB.Timer Timer1 
      Left            =   2880
      Top             =   1680
   End
   Begin VB.CommandButton BotonDel 
      Caption         =   "Delicado"
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton BotonMuy 
      Caption         =   "Muy Sucio"
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   11
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton BotonNor 
      Caption         =   "Normal"
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   10
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton BotonEco 
      Caption         =   "Economico"
      DragIcon        =   "LAVADORA.frx":000C
      Height          =   375
      Index           =   0
      Left            =   0
      MouseIcon       =   "LAVADORA.frx":695E
      TabIndex        =   9
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox TempSelec 
      CausesValidation=   0   'False
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   " "
      Top             =   840
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "sensores"
      Height          =   1335
      Left            =   2040
      TabIndex        =   3
      Top             =   0
      Width           =   1095
      Begin VB.TextBox TempLeida 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.Image Image3 
         Height          =   285
         Left            =   120
         Picture         =   "LAVADORA.frx":D2B0
         Stretch         =   -1  'True
         Top             =   960
         Width           =   405
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   120
         Picture         =   "LAVADORA.frx":D954
         Stretch         =   -1  'True
         Top             =   600
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   315
         Left            =   120
         Picture         =   "LAVADORA.frx":DD12
         Stretch         =   -1  'True
         Top             =   240
         Width           =   285
      End
      Begin VB.Label Label5 
         DragIcon        =   "LAVADORA.frx":DFE7
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.Label MotorLabel 
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   960
         Width           =   375
      End
      Begin VB.Label NivelAgua 
         BackColor       =   &H00FFFF00&
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.TextBox Error_Texto 
      CausesValidation=   0   'False
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox FaseTexto 
      CausesValidation=   0   'False
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton ON_OFF_boton 
      Appearance      =   0  'Flat
      Caption         =   "I/O"
      Height          =   255
      Left            =   960
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conectado As Boolean
Dim control As ControladorEventos

'Establece el lavado delicado
Private Sub BotonDel_Click()
    Timer1.Enabled = False
    control.Delicado
    Timer1.Enabled = True
End Sub
'Establece el lavado economico
Private Sub BotonEco_Click(Index As Integer)
    Timer1.Enabled = False
    control.Economico
    Timer1.Enabled = True
End Sub
'Establece el lavado muy sucio
Private Sub BotonMuy_Click(Index As Integer)
    Timer1.Enabled = False
    control.MuySucio
    Timer1.Enabled = True
End Sub
'Establece el lavado normal
Private Sub BotonNor_Click(Index As Integer)
    Timer1.Enabled = False
    control.Normal
    Timer1.Enabled = True
End Sub


'Procedimiento que inicializa las variables
Private Sub Form_Load()
    Set control = New ControladorEventos
    conectado = False
    control.inicia
    
End Sub


'Boton Encendido
Private Sub ON_OFF_boton_Click()
    If conectado Then
       control.desactivaTodo
       control.OnOffPlaca
       Principal.Timer1.Enabled = False
       conectado = False
    Else
        control.OnOffPlaca
        Principal.Timer1.Enabled = True
        Principal.Timer1.Interval = 200
        conectado = True
    End If
End Sub

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ProgressBar1.Min = 0
    ProgressBar1.Max = 4
    ProgressBar1.Value = 0
    
End Sub

'Funcion TIMER que cada cierto tiempo lee las entradas y reacciona dependiendo de su estado
Private Sub Timer1_Timer()
    'Comprobamos si la entrada de ON/OFF esta conectada para poder reaccionar
    If control.estaON Then
        'Leemos los sensores y actualizamos las variables de éstos
        control.leeSensores
        'Comprobacion si puerta abierta o cerrada
        If control.puertaAbierta Then
            control.gestionaPuertaAbierta
        Else
            'Comprobacion del nivel de carga: media o completa solo una vez por ejecucion del prograna de lavado
            If Not control.gestionCarga Then
                If Not control.modoActivo Then
                    control.gestionarCarga
                End If
            Else
                If control.gestionandoMediaCarga Then
                        control.gestionarMediaCarga
                        
                Else
                    If control.gestionandoCargaCompleta Then
                        control.gestionarCargaCompleta
                    Else
                    
                        'Comprobacion del sobrenivel de agua
                        If control.sobrenivelAgua Then
                            control.gestionarSobrenivelAgua
                        Else
                            ' comprobando si esta todabia vaciando el tambor provocado por el sobrenivel
                            If control.vaciandoTambor Then
                                control.vaciaTamborNivel
                            Else
                                'Comprobacion de la temperatura del termostato
                                If control.termostato Then
                                    control.gestionarTermostato
                                Else
                                    If control.controlandoTemperaturaLavado Then
                                        control.gestionandoTermostato
                                    Else
                                    'Si no existe ningun error, continuamos el lavado
                                    control.gestionaLavado
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        'Si esta apado, desactivamos todas las salidas para asegurarnos
        control.desactivaTodo
    End If
End Sub


