VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Placa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private WithEvents Lip As LipAdq
Attribute Lip.VB_VarHelpID = -1
Private conectada As Boolean
Private motorAct As Boolean
Private centrifugando As Boolean


'###Procedimiento al que se debe llamar inicialmente para inicializar todas las variables necesarias
Public Sub inicializa()
    conectada = False
    Set Lip = New LipAdq
    motorAct = False
    centrifugando = False
End Sub

'### Conectamos o desconectamos la placa ###
Public Sub OnOffPlaca()
    If conectada Then
        conectada = False
        Lip.Desconectar
    Else
        conectada = True
        Lip.ConectarUSB
    End If
End Sub


'=====================================================================
'=====================================================================
'============== LECTURAS A REALIZAR EN LA PLACA ======================
'=====================================================================
'=====================================================================

'### Nos indica si la puerta esta abierta o cerrada ###
Public Function puertaAbierta() As Boolean
    ' Mirar el valor de la salida digital de puerta abierta
    If Lip.EntradaDigital(2) = 1 Then
        puertaAbierta = True
    Else
        puertaAbierta = False
    End If
End Function

'### Nos indica si la lavadora esta a media carga o a carga completa ###
'### VERDADERO si esta a carga completa, FALSO si esta a media carga ###
Public Function cargaCompleta() As Boolean
    'VERDADERO si esta a carga completa, FALSO si esta a media carga
    If Lip.EntradaDigital(3) = 1 Then
        cargaCompleta = False
    Else
        cargaCompleta = True
    End If
End Function


'### Indica si la lavadora esta encendida ###
Public Function estaON() As Boolean
    'Comprueba el valor de la salida digital ON/OFF
    If Lip.EntradaDigital(1) = 1 Then
        estaON = False
    Else
        estaON = True
    End If
End Function


'### Indica si ha saltado el termostato de seguridad ###
Public Function termostato() As Boolean
    'Comprueba el valor de la salida digital de Termostato
     If Lip.EntradaDigital(5) = 1 Then
        termostato = False
    Else
        termostato = True
    End If
End Function


'### Indica si se ha activado el sensor de nivel del agua ###
Public Function sobrenivelAgua() As Boolean
    'Comprueba el valor de la salida digital de Sobrenivel de Agua
     If Lip.EntradaDigital(4) = 1 Then
        sobrenivelAgua = False
    Else
        sobrenivelAgua = True
    End If
End Function


'### Devuelve el valor leido por el potenciómetro de nivel de agua ###
Public Function leeNivelAgua() As Integer
    'Lee el valor del potenciómetro del nivel de agua y lo procesa para indicar un valor con sentido
    leeNivelAgua = Lip.EntradaAnalogica(1) / 40
End Function


'### Devuelve el valor leido por el potenciómetro de temperatura de lavado ###
Public Function leeTempLavado() As Integer
    'Lee el valor del potenciómetro de temperatura de lavado y lo procesa para indicar un valor con sentido
    leeTempLavado = 30 + (Lip.EntradaAnalogica(2) / 30)
End Function


'### Devuelve el valor leido por el NTC de temperatura del agua ###
Public Function leeTempAgua() As Integer
    'Lee el valor del NTC de temperatura del agua y lo procesa para indicar un valor con sentido
    leeTempAgua = Lip.EntradaAnalogica(0) - 440
End Function


'=======================================================================
'=======================================================================
'============== ESCRITURAS A REALIZAR EN LA PLACA ======================
'=======================================================================
'=======================================================================



'###########################  AGUA  ####################################

'### Permite la entrada de Agua a la lavadora ###
Public Sub activaEntradaAgua()
       Lip.SalidaDigital(3) = 1
End Sub

'### Impide la entrada de Agua a la lavadora ###
Public Sub desactivaEntradaAgua()
         Lip.SalidaDigital(3) = 0
End Sub



'###########################  DETERGENTE  ###############################

'### Permite la entrada de Detergente a la lavadora ###
Public Sub activaEntradaDetergente()
     Lip.SalidaDigital(2) = 1
End Sub

'### Impide la entrada de Detergente a la lavadora ###
Public Sub desactivaEntradaDetergente()
    Lip.SalidaDigital(2) = 0
End Sub



'###########################  MOTOR  ###############################

'### Enciende el motor ###
Public Sub activaMotor()
     Lip.SalidasDigitales = 1
     motorAct = True
End Sub

'### Apaga el motor ###
Public Sub desactivaMotor()
     Lip.SalidaDigital(0) = 0
     motorAct = False
End Sub

'### Comprueba si el motor esta encendido
Public Function motorEncendido() As Boolean
    motorEncendido = motorAct
End Function


'###########################  CENTRIFUGADO  #########################

'### Enciende el motor en modo Centrifugado ###
Public Sub activaCentrifugado()
     Lip.SalidaDigital(1) = 1
     Lip.SalidaDigital(0) = 1
     centrifugando = True
     motorAct = True
End Sub

'### Apaga el motor en modo Centrifugado ###
Public Sub desactivaCentrifugado()
     Lip.SalidaDigital(1) = 0
     Lip.SalidaDigital(0) = 0
     centrifugando = False
     motorAct = False
End Sub

'### Comprueba si el motor esta centrifugando
Public Function motorCentrifugando() As Boolean
    motorCentrifugando = centrifugando
End Function


'###########################  DESAGÜE  #############################

'### Activa la bomba de desague ###
Public Sub activaDesague()
     Lip.SalidaDigital(4) = 1
End Sub

'### Apaga la bomba de desague ###
Public Sub desactivaDesague()
     Lip.SalidaDigital(4) = 0
End Sub

'###########################  CALEFACTOR  #############################
'### Activa la calefaccion del agua ###
Public Sub activaCalefaccion()
     Lip.SalidaDigital(5) = 1
End Sub

'### Desactiva la calefaccion del agua ###
Public Sub desactivaCalefaccion()
     Lip.SalidaDigital(5) = 0
End Sub

'### Desactiva todas las salidas ####
Public Sub desactivaTodo()
    Lip.SalidasDigitales = 0
End Sub


