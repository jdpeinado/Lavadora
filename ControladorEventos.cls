VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControladorEventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private pl As Placa
Private tempIntroducida, aguaIntroducida As Integer
Private agua As Double
Private t, tEntradaDetergente, tMotor, tCentrifugado As Integer
Private fase As Integer 'Inactivo=0, Prelavado=1, Lavado=2, Aclarado=3, Centrifugado=4
'Variables para controlar procesos de lavado
Private puedeSeleccionarModo, tamborLleno, motorMovido, tamborVacio, tamborVaciando, DetergenteIntroducido, motorCentrifugado, tempLavadoAlcanzada, gestionTermostato As Boolean
'Variables para controlar las repeticiones de cada proceso del lavado
Private nLavados, LavadosProgramados, nAclarados, AclaradosProgramados, nCentrifugados, CentrifugadosProgramados As Integer
'Variables que controlan la Carga de Lavado: Completa o Media
Private nivelMediaCarga, nivelCargaCompleta As Integer
Private carga, cargaComp, cargaMedia As Boolean



'IMPORTANTE: Es necesario llamar primero a esta función para iniciar todos los parámetros de la clase
Public Sub inicia()
    Set pl = New Placa
    pl.inicializa
    fase = 30
    tEntradaDetergente = 30
    tMotor = 30
    tCentrifugado = 30
    puedeSeleccionarModo = True
    nivelCargaCompleta = 10
    nivelMediaCarga = 5
End Sub


'Procedimiento que lee los sensores y actualiza los campos de la interfaz necesarios
Public Sub leeSensores()
    'Establecemos el texto de error a la cadena vacía
    Principal.Error_Texto.Text = ""
    'Comprobamos si se ha modificado los litros que queremos que entren en el tambor
    If aguaIntroducida <> pl.leeNivelAgua Then
        aguaIntroducida = pl.leeNivelAgua
        Principal.Error_Texto.Text = pl.leeNivelAgua
        Principal.Error_Texto.Text = Principal.Error_Texto.Text + " litros capacidad"
        Pausa 0.5
    End If
        
    'Actualizamos los valores de la temperatura seleccionada y el sensor del interior del tambor
    Principal.TempSelec.Text = pl.leeTempLavado
    Principal.TempSelec.Text = Principal.TempSelec.Text + " º"
    Principal.TempLeida.Text = pl.leeTempAgua
    Principal.TempLeida.Text = Principal.TempLeida.Text + " º"
    Principal.NivelAgua.Caption = agua
    Principal.NivelAgua.Caption = Principal.NivelAgua.Caption + " l."
    
    'Comprobacion si motor encendido o apagado o centrifugando y lo representamos
    If pl.motorEncendido Then
        Principal.MotorLabel.BackColor = &HFF00&
        If pl.motorCentrifugando Then
            Principal.MotorLabel.Caption = " C "
        Else
            Principal.MotorLabel.Caption = " M "
        End If
    Else
        Principal.MotorLabel.BackColor = &H8000000F
        Principal.MotorLabel.Caption = ""
    End If
End Sub



'-----GESTIONAR TODO EL BUCLE DE LAVADO------------------
Public Sub gestionaLavado()
    'Se realiza todo el bucle de lavado
    t = t + 1
    
    
    'Si la fase es 0, la lavadora no esta haciendo nada
    If fase = 0 Then
        Principal.FaseTexto.Text = "Selecciona Modo"
        puedeSeleccionarModo = True
        
    End If
    'Esta en la fase de prelavado
    If fase = 1 Then
        Principal.FaseTexto.Text = "1.- PRELAVADO"
        Prelavado
        Principal.ProgressBar1.Value = fase
    End If
    'Esta en la fase de lavado
    If fase = 2 Then
        Principal.FaseTexto.Text = "2.- LAVADO"
        lavado
        Principal.ProgressBar1.Value = fase
    End If
    'Esta en la fase de aclarado
    If fase = 3 Then
        Principal.FaseTexto.Text = "3.- ACLARADO"
        aclarado
        Principal.ProgressBar1.Value = fase
    End If
    'Esta en la fase de centrifugado
    If fase = 4 Then
        'Si no hay programado ningun centrifugado pasamos a fase 0
        If CentrifugadosProgramados > 0 Then
            Principal.FaseTexto.Text = "4.- CENTRIFUGADO"
            centrifugado
        Else
            fase = 0
            Principal.ProgressBar1.Value = fase
        End If
    End If
End Sub


'##########  FUNCIONES NECESARIAS PARA EL LAVADO  ############################

'-Prelavado
Private Sub Prelavado()
  If tamborLleno Then
    If tempLavadoAlcanzada Then
        If DetergenteIntroducido Then
            If motorMovido Then
                If tamborVacio Then
                    cambiaFase 2
                Else
                    vaciaTambor
                End If
            Else
                MueveMotor
                Principal.Error_Texto.Text = Principal.Error_Texto.Text + " prelavado "
            End If
        Else
          introduceDetergente
        End If
    Else
        calentarAguaLavado
    End If
  Else
    llenaTambor
  End If

End Sub


'-Lavado
Private Sub lavado()
  If tamborLleno Then
     If tempLavadoAlcanzada Then
        If DetergenteIntroducido Then
            If motorMovido Then
                If tamborVacio Then
                    If nLavados < LavadosProgramados Then
                        nLavados = nLavados + 1
                        cambiaFase 2
                    Else
                        cambiaFase 3
                        
                    End If
                 
                Else
                    vaciaTambor
                End If
            Else
                MueveMotor
                Principal.Error_Texto.Text = Principal.Error_Texto.Text + " lavado "
            End If
        Else
          introduceDetergente
        End If
    Else
        calentarAguaLavado
    End If
  Else
    llenaTambor
  End If
    
End Sub

'-Aclarado
Private Sub aclarado()
    If tamborLleno Then
        If motorMovido Then
            If tamborVacio Then
                If nAclarados < AclaradosProgramados Then
                    nAclarados = nAclarados + 1
                    cambiaFase 3
                Else
                    cambiaFase 4
                End If
            Else
                vaciaTambor
            End If
        Else
            MueveMotor
            Principal.Error_Texto.Text = Principal.Error_Texto.Text + " aclarado "
        End If
    Else
        llenaTambor
    End If

End Sub

'-Centrifugado
Private Sub centrifugado()
    If motorCentrifugado Then
        If nCentrifugados < CentrifugadosProgramados Then
            nCentrifugados = nCentrifugados + 1
            cambiaFase 4
        Else
            cambiaFase 0
        End If
    Else
        MueveCentrifugado
    End If
End Sub



'Inicializa las variables necesarias para entrar en la nueva fase del lavado
Private Sub cambiaFase(nuevaFase As Integer)
    tamborLleno = False
    motorMovido = False
    tamborVacio = False
    DetergenteIntroducido = False
    tempLavadoAlcanzada = False
    motorCentrifugado = False
    fase = nuevaFase
    
End Sub

'__Llena el tambor de la lavadora hasta el numero de litros introducido como parametro
Private Sub llenaTambor()
   Principal.Error_Texto.Text = "Llenando Tambor"
   pl.activaEntradaAgua
   agua = agua + 0.5
   If agua >= aguaIntroducida Then
    tamborLleno = True
    pl.desactivaEntradaAgua
    t = 0
   End If
   
   
End Sub

'__Vacia el tambor de la lavadora
Private Sub vaciaTambor()
   
   If agua <= 0 Then
        If puertaAbierta Then
            tamborVacio = False
        Else
            tamborVacio = True
        End If
        pl.desactivaDesague
        t = 0
   Else
       Principal.Error_Texto.Text = "Vaciando Tambor"
       pl.desactivaMotor
       pl.activaDesague
       agua = agua - 0.5
   End If
      
End Sub
Public Function vaciandoTambor() As Boolean
    vaciandoTambor = tamborVaciando

End Function

Public Sub vaciaTamborNivel()
    If agua > pl.leeNivelAgua Then
        Principal.Error_Texto.Text = "Vaciando Tambor"
        pl.desactivaMotor
        pl.activaDesague
        agua = agua - 0.5
    Else
        tamborVaciando = False
        pl.desactivaDesague
    End If
End Sub


'__Activa el motor
Private Sub MueveMotor()
   Principal.Error_Texto.Text = "Procesando "
   pl.activaMotor
   If t >= tMotor Then
        motorMovido = True
        pl.desactivaMotor
        t = 0
   End If
End Sub

'__Activa el centrifugado
Private Sub MueveCentrifugado()
   Principal.Error_Texto.Text = "Motor centrifugando "
   pl.activaCentrifugado
   pl.activaDesague
 '   pl.activaDesague
   If t >= tCentrifugado Then
        motorCentrifugado = True
        pl.desactivaCentrifugado
        pl.desactivaDesague
        t = 0
   End If
End Sub

'__Activa Entrada de detergente
Private Sub introduceDetergente()
   Principal.Error_Texto.Text = "Introduciendo detergente y aditivos "
   pl.activaEntradaDetergente
   If t >= tEntradaDetergente Then
        DetergenteIntroducido = True
        pl.desactivaEntradaDetergente
        t = 0
   End If
End Sub
'__Activa el calefactor para calentar el agua de lavado
Private Sub calentarAguaLavado()
   Principal.Error_Texto.Text = "Calentando agua para lavado "
   pl.activaCalefaccion
   If leeTempAgua >= leeTempLavado Then
        pl.desactivaCalefaccion
        tempLavadoAlcanzada = True
        Principal.Error_Texto.Text = " Temperatura de Lavado Alcanzada"
        Pausa 0.5
   End If
End Sub





'-----GESTIONAR PUERTA ABIERTA-------------------
Public Sub gestionaPuertaAbierta()
    Principal.Error_Texto.Text = "Puerta abierta: Vaciando tambor"
    vaciaTambor
    'Ponemos estas variables a false para que cuando se cierre la puerta, se hagan otra vez
    tamborLleno = False
    DetergenteIntroducido = False
    motorCentrifugado = False
    motorMovido = False
    tempLavadoAlcanzada = False
    
End Sub
' ------- COMPRUEBA QUE TIPO DE CARGA SE HA SELECCIONADO PARA EL LAVADO Y LO GESTIONA
Public Function gestionCarga() As Boolean
    gestionCarga = carga
End Function
Public Sub gestionarCarga()
    If cargaCompleta Then
        cargaComp = True
        Principal.Error_Texto.Text = "Carga completa Seleccionada"
        Pausa 0.5
    Else
        cargaMedia = True
        Principal.Error_Texto.Text = "Media Carga Seleccionada"
        Pausa 0.5
    End If
    carga = True
End Sub

'-----GESTIONAR CARGA COMPLETA-------------------
Public Function gestionandoCargaCompleta() As Boolean
    gestionandoCargaCompleta = cargaComp
End Function

Public Function gestionandoMediaCarga() As Boolean
    gestionandoMediaCarga = cargaMedia
End Function

Public Sub gestionarMediaCarga()
    Principal.Error_Texto.Text = "Llenando Tambor"
    Pausa 0.5
    If leeNivelAgua > nivelMediaCarga Then
        cargaMedia = False
        Principal.Error_Texto.Text = leeNivelAgua
    End If
End Sub

Public Sub gestionarCargaCompleta()
    Principal.Error_Texto.Text = "Llenando Tambor"
    Pausa 0.5
    If leeNivelAgua >= nivelCargaCompleta Then
        cargaComp = False
    End If
End Sub



'-----GESTIONAR SOBRENIVEL AGUA-------------------
Public Sub gestionarSobrenivelAgua()
    'gestionaPuertaAbierta
    Principal.Error_Texto.Text = "Nivel de agua superado"
    agua = 15
    tamborVaciando = True
    vaciaTamborNivel
    
End Sub

'-----GESTIONAR TERMOSTATO-------------------
Public Sub gestionarTermostato()
    Principal.Error_Texto.Text = "Temperatura superada: Desactiva Calefactor"
    pl.desactivaCalefaccion
    pl.desactivaMotor
    pl.desactivaCentrifugado
    If tamborLleno Then
        gestionTermostato = True
    End If
    
End Sub
Public Sub gestionandoTermostato()
            
    If Not tamborVacio Then
        Principal.Error_Texto.Text = "Temperatura superada: vaciando tambor"
        vaciaTambor
        tamborLleno = False
    Else
        Principal.Error_Texto.Text = "Temperatura superada: llenando tambor"
        llenaTambor
        If tamborLleno Then
            gestionTermostato = False
            tamborVacio = False
            
        End If
    End If
End Sub
Public Function controlandoTemperaturaLavado() As Boolean
    controlandoTemperaturaLavado = gestionTermostato
End Function
'-----DESACTIVA TODAS LAS SALIDAS-------------------
Public Sub desactivaTodo()
    pl.desactivaTodo
End Sub

'-----ENCIENDE O APAGA LA PLACA-------------------
Public Sub OnOffPlaca()
    pl.OnOffPlaca
End Sub

'######################## MODOS DE LAVADO ###################################
Public Sub Economico()
    If puedeSeleccionarModo Then
        Principal.Error_Texto.Text = "Lavado económico activado"
        inicializaVariablesLavado False, 1, 1, 1
        Pausa 0.5
        puedeSeleccionarModo = False
    End If
End Sub

Public Sub Normal()
    If puedeSeleccionarModo Then
        Principal.Error_Texto.Text = "Lavado normal activado"
        inicializaVariablesLavado False, 2, 2, 1
        Pausa 0.5
        puedeSeleccionarModo = False
    End If
End Sub

Public Sub Delicado()
    If puedeSeleccionarModo Then
        Principal.Error_Texto.Text = "Lavado delicado activado"
        inicializaVariablesLavado False, 1, 1, 0
        Pausa 0.5
        puedeSeleccionarModo = False
    End If
End Sub

'Establece el programa de lavamo muy sucio
Public Sub MuySucio()
    If puedeSeleccionarModo Then
        Principal.Error_Texto.Text = "Lavado muy sucio activado"
        inicializaVariablesLavado True, 2, 3, 2
        Pausa 0.5
        puedeSeleccionarModo = False
    End If
End Sub

'---Inicializa las variables que intervienen en el lavado en el momento de elegir el modo de lavado
Private Sub inicializaVariablesLavado(Prelavado As Boolean, Lavados As Integer, Aclarados As Integer, Centrifugados As Integer)
    nLavados = 1
    LavadosProgramados = Lavados
    nAclarados = 1
    AclaradosProgramados = Aclarados
    nCentrifugados = 1
    CentrifugadosProgramados = Centrifugados
    tamborLleno = False
    motorMovido = False
    tamborVacio = False
    tamborVaciando = False
    gestionTermostato = False
    DetergenteIntroducido = False
    tempLavadoAlcanzada = False
    motorCentrifugado = False
    carga = False
    cargaComp = False
    cargaMedia = False
    t = 0
    'Si indicamos que no hay prelavado, el programa se inicia en la fase 2, si hay en la 1
    If Prelavado Then
        fase = 1
    Else
        fase = 2
    End If
    agua = 0
End Sub


'######################### CONSULTORES #######################################

'### nos indica si el programa esta seleccionado
Public Function modoActivo() As Boolean
    modoActivo = puedeSeleccionarModo
End Function
'### Nos indica si la puerta esta abierta o cerrada ###
Public Function puertaAbierta() As Boolean
    puertaAbierta = pl.puertaAbierta
End Function

'### Nos indica si la lavadora esta a media carga o a carga completa ###
'### VERDADERO si esta a carga completa, FALSO si esta a media carga ###
Public Function cargaCompleta() As Boolean
    cargaCompleta = pl.cargaCompleta
End Function

'### Indica si la lavadora esta encendida ###
Public Function estaON() As Boolean
    estaON = pl.estaON
End Function

'### Indica si ha saltado el termostato de seguridad ###
Public Function termostato() As Boolean
    termostato = pl.termostato
End Function

'### Indica si se ha activado el sensor de nivel del agua ###
Public Function sobrenivelAgua() As Boolean
    sobrenivelAgua = pl.sobrenivelAgua
End Function

'### Devuelve el valor leido por el potenciómetro de nivel de agua ###
Public Function leeNivelAgua() As Integer
   leeNivelAgua = pl.leeNivelAgua
End Function

'### Devuelve el valor leido por el potenciómetro de temperatura de lavado ###
Public Function leeTempLavado() As Integer
    leeTempLavado = pl.leeTempLavado
End Function


'### Devuelve el valor leido por el NTC de temperatura del agua ###
Public Function leeTempAgua() As Integer
    leeTempAgua = pl.leeTempAgua
End Function

'### Comprueba si el motor esta centrifugando
Public Function motorCentrifugando() As Boolean
    motorCentrifugando = pl.motorCentrifugando
End Function

'### Comprueba si el motor esta encendido
Public Function motorEncendido() As Boolean
    motorEncendido = pl.motorEncendido
End Function


'Procedimiento hace parar unos segundos la aplicación
Sub Pausa(Segundos As Double)
    Dim inicio As Double
    ' Devuelve la cantidad de segundos desde que inicio windows
    inicio = Timer
    Do While (Timer - inicio) < Segundos
        DoEvents
    Loop
End Sub

