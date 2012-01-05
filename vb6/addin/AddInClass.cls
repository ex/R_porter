VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddInClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Implements IDTExtensibility


Public WithEvents PrjHandler  As VBProjectsEvents          'controlador de eventos de proyectos
Attribute PrjHandler.VB_VarHelpID = -1
Public WithEvents CmpHandler  As VBComponentsEvents        'controlador de eventos de componentes
Attribute CmpHandler.VB_VarHelpID = -1
Public WithEvents CtlHandler  As VBControlsEvents          'controlador de eventos de controles
Attribute CtlHandler.VB_VarHelpID = -1
Public WithEvents MenuHandler As CommandBarEvents          'controlador de eventos de barra de comandos
Attribute MenuHandler.VB_VarHelpID = -1

Dim mcbMenuCommandBar         As Office.CommandBarControl  'objeto barra de comandos

'------------------------------------------------------------------------------------------
' CLASE
'------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    
    On Error Resume Next
    
    If frmConfiguracion.Visible = True Then
        Unload frmConfiguracion
        Set frmConfiguracion = Nothing
    End If
    
    If frmReporte.Visible = True Then
        Unload frmReporte
        Set frmReporte = Nothing
    End If
    
End Sub

Private Sub Class_Initialize()

    If GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\=ex=\exAddInTable\1.0", "path", gs_Path) = False Then
        MsgBox "No se encuentra ruta de plantillas" & vbCrLf & "Prueba con instalar de nuevo el add-in", vbExclamation, "Error"
        gb_PathFind = False
    Else
        gb_PathFind = True
    End If
    
End Sub

'------------------------------------------------------------------------------------------
' EVENTOS IDTExtensibility
'------------------------------------------------------------------------------------------

'------------------------------------------------------------
'este método agrega el complemento al menú Herramientas de VB
'lo llama el Administrador de complementos de VB
'------------------------------------------------------------
Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As VBIDE.vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
    
    On Error GoTo Handler
    
    Dim aiTmp As AddIn
    
    'guardar la instancia de vb
    Set gVBInstance = VBInst
    
    If Not gwinWindow Is Nothing Then
        'ya se está ejecutando; sólo se muestra
        Show
        If ConnectMode = vbext_cm_AfterStartup Then
            'iniciado desde el Administrador de complementos
            AddToCommandBar
        End If
        Exit Sub
    End If
    
    'crear la ventana Herramientas
    If ConnectMode = vbext_cm_External Then
        
        'hay que ver si ya se está ejecutando
        On Error Resume Next
        Set aiTmp = gVBInstance.Addins("exAddInClass.AddInClass")
        
        On Error GoTo Handler
        
        If aiTmp Is Nothing Then
            'la aplicación no está en el archivo VBADDIN.INI por lo que no está en la colección
            'intentemos usar el primer complemento de la colección
            'para ejecutar esta aplicación y, si no hay ninguno, se producirá
            'un error y esta aplicación no se ejecutará
            Set gwinWindow = gVBInstance.Windows.CreateToolWindow(gVBInstance.Addins(1), "exAddInTAble.usrdoc", "exAddIn", guidMYTOOL$, gdoc_usrdoc)
        Else
            If aiTmp.Connect = False Then
                Set gwinWindow = gVBInstance.Windows.CreateToolWindow(aiTmp, "exAddInTAble.usrdoc", "exAddIn", guidMYTOOL$, gdoc_usrdoc)
            End If
        End If
    Else
        'debe llamarse desde el Administrador de complementos
        Set gwinWindow = gVBInstance.Windows.CreateToolWindow(AddInInst, "exAddInTAble.usrdoc", "exAddIn", guidMYTOOL$, gdoc_usrdoc)
    End If
    
    'recibe el controlador de eventos de proyecto, componentes y controles
    Set Me.PrjHandler = gVBInstance.Events.VBProjectsEvents
    Set Me.CmpHandler = gVBInstance.Events.VBComponentsEvents(Nothing)
    Set Me.CtlHandler = gVBInstance.Events.VBControlsEvents(Nothing, Nothing)
    
    If ConnectMode = vbext_cm_External Then
        'se inicia desde la barra de herramientas de complementos
        Show
    ElseIf ConnectMode = vbext_cm_AfterStartup Then
        'se inicia desde el Adminstrador de complementos
        AddToCommandBar
    End If
    
    Exit Sub
  
Handler:

    MsgBox Err.Description, vbCritical, "OnConnection()"
End Sub

'------------------------------------------------------
'este evento quita el menú barra de comandos
'se llama desde el Administrador de complementos de VB
'------------------------------------------------------
Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As VBIDE.vbext_DisconnectMode, custom() As Variant)
    
    On Error GoTo Handler
    
    'eliminar la entrada barra de comandos
    mcbMenuCommandBar.Delete
    
    Set gwinWindow = Nothing
    
Handler:

End Sub

'------------------------------------------------------------------------------------------
'este evento se desencadena cuando se carga completamente el IDE
'------------------------------------------------------------------------------------------
Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    AddToCommandBar
End Sub

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
' Comentario para impedir que se elimine
' el procedimiento en la compilación.
End Sub

'------------------------------------------------------------------------------------------
' EVENTOS MenuHandler
'------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------
'este evento se desencadena cuando se hace clic en el control barra de comandos en el IDE
'------------------------------------------------------------------------------------------
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Show
End Sub

'------------------------------------------------------------------------------------------
' FUNCIONES AUXILIARES
'------------------------------------------------------------------------------------------
Sub AddToCommandBar()
    
    On Error GoTo Handler
    
    'asegurarse de que la barra de herramientas estándar está visible
    gVBInstance.CommandBars(2).Visible = True
    
    'agregar la barra de comandos
    'la línea siguiente agregará el administrador de TabOrder a la barra de
    'herramientas estándar, situada a la derecha del botón Cuadro de herramientas
    Set mcbMenuCommandBar = gVBInstance.CommandBars(2).Controls.Add(1, , , gVBInstance.CommandBars(2).Controls.Count)
    'establecer el título
    mcbMenuCommandBar.Caption = "exAddInTable"
    
    'copiar el icono al Portapapeles
    Clipboard.SetData LoadResPicture(103, vbResBitmap), vbCFBitmap
    'establecer el icono para el botón
    mcbMenuCommandBar.PasteFace
    
    'recibir el evento
    Set Me.MenuHandler = gVBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    
    Exit Sub
    
Handler:
    MsgBox Err.Description, vbCritical, "AddToCommandBar"
    
End Sub


