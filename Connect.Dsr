VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   7425
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   10065
   _ExtentX        =   17754
   _ExtentY        =   13097
   _Version        =   393216
   Description     =   "The ultimate solution to the Select Case [Enum] problem."
   DisplayName     =   "grigri's Select Case Enum"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo ERROR_HANDLER
    
    'save the vb instance
    Set VBInstance = Application
    
    If Not Compiled Then
        MsgBox "Sorry, this add-in has to be compiled to work." & vbCrLf & _
               "It uses windows hooks, which have to run in the same thread" & vbCrLf & _
               "as the calling application, which only happens when the add-in is compiled", vbCritical
        Exit Sub
    End If
    
    Running = True
    If Running = False Then
        MsgBox "Failed to initialize", vbCritical
    End If
    
    Exit Sub
    
ERROR_HANDLER:
    
    MsgBox Err.Description
    
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    Running = False
End Sub

Private Function Compiled() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    If Err Then
        Compiled = False
    Else
        Compiled = True
    End If
End Function
