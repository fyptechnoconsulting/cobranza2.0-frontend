VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_GenBod 
   Caption         =   "Generacion de Archivo del BOD"
   ClientHeight    =   6330
   ClientLeft      =   555
   ClientTop       =   2415
   ClientWidth     =   18435
   Icon            =   "Frm_GenBod.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   18435
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   120
      TabIndex        =   21
      Top             =   5640
      Width           =   18135
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   5040
      Width           =   18135
   End
   Begin VB.CommandButton btn_salir 
      Caption         =   "&Salir"
      Height          =   855
      Left            =   5280
      Picture         =   "Frm_GenBod.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton btn_ejecutar 
      Caption         =   "&Ejecutar"
      Height          =   855
      Left            =   3360
      Picture         =   "Frm_GenBod.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3960
      Width           =   1095
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   126812161
      CurrentDate     =   43919
   End
   Begin VB.Frame Frame2 
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   3360
      TabIndex        =   4
      Top             =   240
      Width           =   3015
      Begin VB.OptionButton OP_1500 
         Caption         =   "15 Unico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   19
         Top             =   1560
         Width           =   1335
      End
      Begin VB.OptionButton OP_0030 
         Caption         =   "30 Unico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   18
         Top             =   2280
         Width           =   1215
      End
      Begin VB.OptionButton OP_0027 
         Caption         =   "27 Unico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   17
         Top             =   1200
         Width           =   1335
      End
      Begin VB.OptionButton OP_1300 
         Caption         =   "13 Unico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   16
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton OP_0025 
         Caption         =   "25 Unico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   2280
         Width           =   1335
      End
      Begin VB.OptionButton OP_1000 
         Caption         =   "10 Unico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   1335
      End
      Begin VB.OptionButton OP_0022 
         Caption         =   "22 Unico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   1335
      End
      Begin VB.OptionButton op_0800 
         Caption         =   "8 Unico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton OP_1530 
         Caption         =   "15 y 30"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   8
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton OP_1327 
         Caption         =   "13 y 27"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton OP_0822 
         Caption         =   "8 y 22"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton OP_1025 
         Caption         =   "10 y 25"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   1215
      End
   End
   Begin MSComCtl2.DTPicker DTP_FechaEnvio 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   3600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   126812161
      CurrentDate     =   43919
   End
   Begin MSComCtl2.DTPicker DTP_FechaCobro 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   4440
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   126812161
      CurrentDate     =   43919
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha de Cobro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Lbl_FechaEnvio 
      Caption         =   "Fecha de Envio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
   End
End
Attribute VB_Name = "Frm_GenBod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_ejecutar_Click()
'DTP_FechaEnvio.MinDate = Now
'DTP_FechaCobro.MinDate = Now
Dim a As Variant
Dim b As Variant
Dim c As Variant
Dim d As Variant

c = CStr(Format(DTP_FechaCobro.Value, "DD-MM-YYYY"))
b = CStr(Format(DTP_FechaEnvio.Value, "DD-MM-YYYY"))
a = "D:\fyptechnoconsulting\herramientas\pdi-ce-8.1.0.0-365\kitchen.bat -file=" & """" & "D:\fyptechnoconsulting\fuentes\prod\generaArchivoBOD.kjb" & """" & " -level=basic " & """" & "'" & "0116" & "'" & """" & " " & """" & "'" & b & "'" & """" & " " & """" & "'" & c & "'" & """" & " " & "0" & " " & """" & "'" & "570004393627" & "'" & """" & " " & "'" & "bod15y30" & "'"
Text1.Text = a
d = Shell(a, vbNormalFocus)
Text1.Text = a
Text2.Text = c
'D:\fyptechnoconsulting\fuentes\prod\generaArchivoBOD.kjb" -level=basic "'0116'" "'28-04-2019'" "'30-04-2019'" 0 "'570004393627'" 'bod15y30'


a = Shell("Calc.exe", vbNormalFocus)



End Sub

Private Sub btn_salir_Click()
Unload Me

End Sub

Private Sub Form_Load()
DTP_FechaEnvio.Value = Now
DTP_FechaCobro.Value = Now
MonthView1.Value = Now
'DTP_FechaEnvio.MinDate = Now
'DTP_FechaCobro.MinDate = Now

End Sub
