VERSION 5.00
Begin VB.Form Frm_Menu 
   Caption         =   "Sistema de Generacion de Archivos de Cobro"
   ClientHeight    =   4860
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu Mnu_proceso 
      Caption         =   "&Proceso"
      Begin VB.Menu Mnu_PBod 
         Caption         =   "BOD"
         Begin VB.Menu Mnu_PBGenerarArchivo 
            Caption         =   "Generar Archivo de Cobro"
         End
         Begin VB.Menu Mnu_PBAgregarConvenio 
            Caption         =   "Agregar Convenio"
         End
      End
   End
   Begin VB.Menu Mnu_Ayuda 
      Caption         =   "A&yuda"
   End
   Begin VB.Menu Mnu_Salir 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "Frm_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
