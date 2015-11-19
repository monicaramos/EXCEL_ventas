VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmVentas 
   Caption         =   "Generacion archivo Albaranes Venta"
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7410
   Icon            =   "frmVentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7260
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameEscribir 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7395
      Begin ComctlLib.ProgressBar Pb1 
         Height          =   195
         Left            =   90
         TabIndex        =   22
         Top             =   1170
         Visible         =   0   'False
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   690
         Width           =   7155
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   4920
         TabIndex        =   2
         Top             =   1800
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   1
         Left            =   6120
         TabIndex        =   1
         Top             =   1800
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Fichero generado"
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   450
         Width           =   1395
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   1530
         Picture         =   "frmVentas.frx":1782
         Top             =   450
         Width           =   240
      End
   End
   Begin VB.Frame FrameImportar 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7335
      Begin VB.CommandButton Command3 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5760
         TabIndex        =   9
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Importar"
         Height          =   375
         Left            =   4500
         TabIndex        =   8
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   690
         Width           =   6735
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   450
         Width           =   615
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   900
         Picture         =   "frmVentas.frx":1884
         Top             =   420
         Width           =   240
      End
   End
   Begin VB.Frame FrameConfig 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7335
      Begin VB.TextBox Text8 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   2790
         PasswordChar    =   "*"
         TabIndex        =   13
         Text            =   "Text8"
         Top             =   1290
         Width           =   1485
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Index           =   2
         Left            =   2790
         TabIndex        =   12
         Text            =   "Text8"
         Top             =   930
         Width           =   1485
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Index           =   1
         Left            =   2790
         TabIndex        =   11
         Text            =   "Text8"
         Top             =   570
         Width           =   1515
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Index           =   0
         Left            =   2790
         TabIndex        =   15
         Text            =   "Text8"
         Top             =   1650
         Width           =   1485
      End
      Begin VB.CommandButton cmdConfig 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   6120
         TabIndex        =   19
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdConfig 
         Caption         =   "Guardar"
         Height          =   375
         Index           =   0
         Left            =   4800
         TabIndex        =   17
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Server"
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
         Index           =   8
         Left            =   240
         TabIndex        =   21
         Top             =   600
         Width           =   570
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
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
         Index           =   3
         Left            =   240
         TabIndex        =   20
         Top             =   990
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nro.Máximo de Calidades"
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
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   1710
         Width           =   2145
      End
      Begin VB.Label Label7 
         Caption         =   "CLASIFICACION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   240
         Width           =   2010
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   12
         Left            =   240
         TabIndex        =   14
         Top             =   1350
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Private WithEvents frmC As frmCal
Private NoEncontrados As String




Dim Sql As String
Dim VariasEntradas As String


Dim Albaran As Long
Dim FecAlbaran As String
Dim Socio As String
Dim Campo As String
Dim Variedad As String
Dim TipoEntr As String
Dim KilosNet As String
Dim Calidad(20) As String

Private WithEvents frmMens As frmMensajes 'Registros que no ha entrado con error
Attribute frmMens.VB_VarHelpID = -1



Private Sub cmdConfig_Click(Index As Integer)
Dim I As Integer

    If Index = 1 Then
        Unload Me
    Else
        Sql = ""
        For I = 0 To Text8.Count - 1
            If Text8(I).Text = "" Then Sql = Sql & "Campo: " & I & vbCrLf
        Next I
        If Sql <> "" Then
            Sql = "No pueden haber campos vacios: " & vbCrLf & vbCrLf & Sql
            MsgBox Sql, vbExclamation
            Text8(0).SetFocus
            Exit Sub
        End If
        
        mConfig.MaxCalidades = Text8(0).Text
        mConfig.SERVER = Text8(1).Text
        mConfig.User = Text8(2).Text
        mConfig.password = Text8(3).Text
        
        mConfig.Guardar
        
        vConfiguracion False
'        If varConfig.Grabar = 0 Then End
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim Rc As Byte
Dim Mens As String

    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
        
    If Text2.Text <> "" Then
        If Dir(Text2.Text) <> "" Then
            MsgBox "Fichero ya existe", vbExclamation
            Exit Sub
        Else
            FileCopy App.Path & "\" & mConfig.Plantilla, Text2.Text
            NombreHoja = Text2.Text
        End If
    End If
    
    'Abrimos excel
    Rc = AbrirEXCEL
    
    If Rc = 0 Then
    
        'Si queremos que se vea descomentamos  esto
        MiXL.Application.visible = False
'        MiXL.Parent.Windows(1).Visible = False
    
        'Realizamos todos los datos
        'abrimos conexion
        If AbrirConexion(BaseDatos) Then
        
            Screen.MousePointer = vbHourglass
            
            'Vamos linea a linea
            Mens = "Error insertando en Excel"
            ' en el caso de que sea catadau la salida es diferente
            If Cooperativa = 0 Then
                If Not RecorremosLineasCatadau(Mens) Then
                    MsgBox Mens, vbExclamation
                End If
            
            Else
                
                If Not RecorremosLineas(Mens) Then
                    MsgBox Mens, vbExclamation
                End If
                
            End If
            Screen.MousePointer = vbDefault
            
        End If
    
        'Cerramos el excel
        CerrarExcel
                
        MsgBox "Proceso finalizado", vbExclamation
        
        Command1_Click (1)

    End If
    
End Sub

Private Sub Command2_Click()
Dim Rc As Byte
Dim I As Integer
Dim Rs1 As ADODB.Recordset
Dim KilosI As Long


    'IMPORTAR


    If Text5.Text = "" Then
        MsgBox "Escriba el nombre del fichero excel", vbExclamation
        Exit Sub
    End If
        
        
    If Dir(Text5.Text) = "" Then
        MsgBox "Fichero no existe"
        Exit Sub
    End If
    
    NombreHoja = Text5.Text
    'Abrimos excel
    Rc = AbrirEXCEL
    
    If Rc = 0 Then
    
        'Realizamos todos los datos
        'abrimos conexion
        If AbrirConexion(BaseDatos) Then
        
            
            'Vamos linea a linea, buscamos su trabajador
            RecorremosLineasLiquidacion
            
        End If
    
        'Cerramos el excel
        CerrarExcel
      


        Dim RS As ADODB.Recordset
        Dim C As Long
        Dim cad As String
        Sql = "Select * from tmpexcel WHERE situacion <> 0 and codusu = " & Usuario


        Set RS = New ADODB.Recordset
        RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Sql = ""
        C = 0
        While Not RS.EOF
            Sql = Sql & (RS!numalbar) & "        "
            If (C Mod 6) = 5 Then Sql = Sql & vbCrLf
            C = C + 1
            RS.MoveNext
        Wend
        RS.Close
        If C > 0 Then
            Set frmMens = New frmMensajes
            
            frmMens.Cadena = "select * from tmpexcel where situacion <> 0 and codusu = " & Usuario
            frmMens.OpcionMensaje = 1
            frmMens.Show vbModal
            
'            SQL = "Se han encontrado " & C & " registros con datos incorrectos en la BD: " & vbCrLf & SQL
'            SQL = SQL & " ¿Desea continuar ?"
'            If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbNo Then Exit Sub
        End If

        'Abrimos los registros =0 k son los OK'
        Sql = "¿ Desea importar las clasificaciones correctas ?"
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub


        Sql = "Select * from tmpexcel WHERE situacion = 0 and codusu = " & Usuario
        
        RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        C = 0
        While Not RS.EOF
            C = C + 1
            
            Sql = "delete from rhisfruta_clasif where numalbar = " & RS!numalbar
            Conn.Execute Sql
            
            
            For I = 1 To mConfig.MaxCalidades
                Select Case I
                    Case 1
                        KilosI = RS!calidad1
                    Case 2
                        KilosI = RS!calidad2
                    Case 3
                        KilosI = RS!calidad3
                    Case 4
                        KilosI = RS!calidad4
                    Case 5
                        KilosI = RS!calidad5
                    Case 6
                        KilosI = RS!calidad6
                    Case 7
                        KilosI = RS!calidad7
                    Case 8
                        KilosI = RS!calidad8
                    Case 9
                        KilosI = RS!calidad9
                    Case 10
                        KilosI = RS!calidad10
                    Case 11
                        KilosI = RS!calidad11
                    Case 12
                        KilosI = RS!calidad12
                    Case 13
                        KilosI = RS!calidad13
                    Case 14
                        KilosI = RS!calidad14
                    Case 15
                        KilosI = RS!calidad15
                    Case 16
                        KilosI = RS!calidad16
                    Case 17
                        KilosI = RS!calidad17
                    Case 18
                        KilosI = RS!calidad18
                    Case 19
                        KilosI = RS!calidad19
                    Case 20
                        KilosI = RS!calidad20
                End Select
                
                If KilosI <> 0 Then
                    Sql = "insert into rhisfruta_clasif (numalbar, codvarie, codcalid, kilosnet) "
                    Sql = Sql & " values (" & RS!numalbar & "," & RS!codvarie & ","
                    Sql = Sql & I & ","
                    Sql = Sql & KilosI & ")"
                
                    Conn.Execute Sql
                End If
                    
            Next I
            
            RS.MoveNext
        Wend
        RS.Close
    End If
    MsgBox "FIN", vbInformation
    
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
'    Combo1.ListIndex = Month(Now) - 1
'    Text3.Text = Year(Now)
    FrameEscribir.visible = False
    FrameImportar.visible = False
    Me.FrameConfig.visible = False
    Limpiar
    Select Case EsImportaci
    Case 0
        Caption = "CONFIGURACION"
        FrameConfig.visible = True
'        vConfiguracion True
    Case 1
        Caption = "Cargar Ventas desde fichero excel"
        FrameImportar.visible = True
    Case 2
        Caption = "Crear fichero Albaranes Ventas"
        FrameEscribir.visible = True
    End Select
    
    
 
End Sub

Private Sub Limpiar()
Dim T As Control
    For Each T In Me.Controls
        If TypeOf T Is TextBox Then
            T.Text = ""
        End If
    Next
        
End Sub
Private Function TransformaComasPuntos(Cadena) As String
Dim cad As String
Dim j As Integer
    
    j = InStr(1, Cadena, ",")
    If j > 0 Then
        cad = Mid(Cadena, 1, j - 1) & "." & Mid(Cadena, j + 1)
    Else
        cad = Cadena
    End If
    TransformaComasPuntos = cad
End Function

Private Sub frmC_Selec(vFecha As Date)
'    Text4.Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image1_Click()
    AbrirDialogo 0
End Sub

Private Sub Image2_Click()
    AbrirDialogo 1
End Sub


Private Sub AbrirDialogo(Opcion As Integer)

    On Error GoTo EA
    
    With Me.CommonDialog1
        Select Case Opcion
        Case 0, 2
            .DialogTitle = "Archivo origen de datos"
        Case 1
            .DialogTitle = "Archivo destino de datos"
        End Select
        .Filter = "EXCEL (*.xls)|*.xls"
        .CancelError = True
        If Opcion <> 1 Then
            .ShowOpen
            If Opcion = 0 Then
                Text2.Text = .FileName
            Else
                Text5.Text = .FileName
            End If
        Else
            .ShowSave
            Text2.Text = .FileName
        End If
        
        
        
    End With
EA:
End Sub

Public Sub IncrementarProgresNew(ByRef PBar As ProgressBar, Veces As Integer)
On Error Resume Next
'    PBar.Value = PBar.Value + ((Veces * PBar.Max) / CInt(PBar.Tag))
    PBar.Value = PBar.Value + Veces
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Function RecorremosLineas(Mens As String) As Boolean
Dim I As Integer
Dim j As Integer
Dim JJ As Integer
Dim F As Date
Dim Cod As String
Dim FE As String
Dim RT As ADODB.Recordset
Dim RS As ADODB.Recordset
Dim RsCAlb As ADODB.Recordset
Dim RsLAlb As ADODB.Recordset
Dim Calidad As Integer
Dim NFic As Integer
Dim Lin As String
Dim Sql1 As String
Dim Sql2 As String

Dim NFile As Integer
Dim TotalGastos As Currency
Dim GastosKg As Currency
Dim ImpVtaKg As Currency
Dim ValorFruta As Currency

    On Error GoTo eRecorremosLineas

    RecorremosLineas = False


    Sql = "select * from tmpinfventas where codusu = " & Usuario & " order by fecalbar, numalbar "
    Sql1 = "select count(*) from tmpinfventas where codusu = " & Usuario

    Set RT = New ADODB.Recordset
    RT.Open Sql1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RT.EOF Then
        Me.Pb1.visible = True
        Me.Pb1.Max = RT.Fields(0).Value
        Me.Pb1.Value = 0
        Me.Refresh
    End If
    
    Set RT = Nothing
    Set RT = New ADODB.Recordset
    RT.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    I = 1
    While Not RT.EOF
        I = I + 1
            
        IncrementarProgresNew Pb1, 1
    
        Sql2 = "select * from albaran where numalbar = " & DBSet(RT!numalbar, "N")
        Set RsCAlb = New ADODB.Recordset
        RsCAlb.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not RsCAlb.EOF Then
    
            Sql2 = "select * from albaran_variedad where numalbar = " & DBSet(RT!numalbar, "N") & " and numlinea = " & DBSet(RT!numlinea, "N")
            Set RsLAlb = New ADODB.Recordset
            RsLAlb.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
            If Not RsLAlb.EOF Then
                'variedad
                ExcelSheet.Cells(I, 1).Value = DBLet(RsLAlb!codvarie, "N") ' codvarie
                
                Sql = "select nomvarie from variedades where codvarie = " & DBSet(RsLAlb!codvarie, "N")
                Set RS = New ADODB.Recordset
                RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS.EOF Then
                    ExcelSheet.Cells(I, 2).Value = RS.Fields(0).Value ' nombre de variedad
                Else
                    ExcelSheet.Cells(I, 2).Value = "" ' nombre de variedad
                End If
                Set RS = Nothing
                
                ' variedad comercial
                ExcelSheet.Cells(I, 3).Value = DBLet(RsLAlb!codvarco, "N") ' codvarie comercial
                
                Sql = "select nomvarie from variedades where codvarie = " & DBSet(RsLAlb!codvarco, "N")
                Set RS = New ADODB.Recordset
                RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS.EOF Then
                    ExcelSheet.Cells(I, 4).Value = RS.Fields(0).Value ' nombre de variedad
                Else
                    ExcelSheet.Cells(I, 4).Value = "" ' nombre de variedad
                End If
                Set RS = Nothing
                
                ' fecalbar
                ExcelSheet.Cells(I, 5).Value = Format(RT!fecalbar, "yyyy/mm/dd")
                ' numalbar
                ExcelSheet.Cells(I, 6).Value = RT!numalbar
                ' numlinea
                ExcelSheet.Cells(I, 7).Value = RT!numlinea
                
                ' cliente
                ExcelSheet.Cells(I, 8).Value = DBLet(RsCAlb!codclien, "N")
                
                Sql = "select nomclien from clientes where codclien = " & DBSet(RsCAlb!codclien, "N")
                Set RS = New ADODB.Recordset
                RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS.EOF Then
                    ExcelSheet.Cells(I, 9).Value = RS.Fields(0).Value
                Else
                    ExcelSheet.Cells(I, 9).Value = ""
                End If
                Set RS = Nothing
                
                ' destino
                ExcelSheet.Cells(I, 10).Value = DBLet(RsCAlb!codDESTI, "N")
                
                Sql = "select nomdesti from destinos where codclien = " & DBSet(RsCAlb!codclien, "N") & " and coddesti = " & DBSet(RsCAlb!codDESTI, "N")
                Set RS = New ADODB.Recordset
                RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS.EOF Then
                    ExcelSheet.Cells(I, 11).Value = RS.Fields(0).Value
                Else
                    ExcelSheet.Cells(I, 11).Value = ""
                End If
                Set RS = Nothing
                
                ' forfait
                ExcelSheet.Cells(I, 12).Value = DBLet(RsLAlb!codforfait, "T")
                
                Sql = "select nomconfe from forfaits where codforfait = " & DBSet(RsLAlb!codforfait, "T")
                Set RS = New ADODB.Recordset
                RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS.EOF Then
                    ExcelSheet.Cells(I, 13).Value = RS.Fields(0).Value
                Else
                    ExcelSheet.Cells(I, 13).Value = ""
                End If
                Set RS = Nothing
                
                ' marca
                ExcelSheet.Cells(I, 14).Value = DBLet(RsLAlb!codmarca, "N")
                
                Sql = "select nommarca from marcas where codmarca = " & DBSet(RsLAlb!codmarca, "N")
                Set RS = New ADODB.Recordset
                RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS.EOF Then
                    ExcelSheet.Cells(I, 15).Value = RS.Fields(0).Value
                Else
                    ExcelSheet.Cells(I, 15).Value = ""
                End If
                Set RS = Nothing
                
                ' categoria
                ExcelSheet.Cells(I, 16).Value = DBLet(RsLAlb!categori, "T")
                
                ' numcajas
                ExcelSheet.Cells(I, 17).Value = DBLet(RT!numcajas, "N")
                ' pesoreal
                ExcelSheet.Cells(I, 18).Value = DBLet(RT!pesoreal, "N")
                ' pesoneto
                ExcelSheet.Cells(I, 19).Value = DBLet(RT!pesoneto, "N")
                ' gastos portes
                ExcelSheet.Cells(I, 20).Value = DBLet(RT!gastosportes, "N")
                ' comisiones
                ExcelSheet.Cells(I, 21).Value = DBLet(RT!gastoscomisiones, "N")
                
                'Gastos Envases
                ExcelSheet.Cells(I, 22).Value = DBLet(RT!gastosenvases, "N")
                                
                ' gastos1
                ExcelSheet.Cells(I, 23).Value = DBLet(RT!gastos1, "N")
                ' gastos2
                ExcelSheet.Cells(I, 24).Value = DBLet(RT!gastos2, "N")
                ' gastos3
                ExcelSheet.Cells(I, 25).Value = DBLet(RT!gastos3, "N")
                ' gastos4
                ExcelSheet.Cells(I, 26).Value = DBLet(RT!gastos4, "N")
                ' gastos5
                ExcelSheet.Cells(I, 27).Value = DBLet(RT!gastos5, "N")
                'total gastos
                TotalGastos = DBLet(RT!gastos1, "N") + DBLet(RT!gastos2, "N") + DBLet(RT!gastos3, "N") + DBLet(RT!gastos4, "N") + DBLet(RT!gastos5, "N") + _
                                                DBLet(RT!gastosportes, "N") + DBLet(RT!gastoscomisiones, "N") + DBLet(RT!gastosenvases, "N")
'                TotalGastos = DBLet(RT!gastos, "N")
                ExcelSheet.Cells(I, 28).Value = TotalGastos
                'gastos/kg
                GastosKg = 0
                If DBLet(RT!pesoneto, "N") <> 0 Then GastosKg = Round2(TotalGastos / DBLet(RT!pesoneto, "N"), 4)
                ExcelSheet.Cells(I, 29).Value = TransformaComasPuntos(GastosKg)
                'importe de venta
                ExcelSheet.Cells(I, 30).Value = DBLet(RT!impventa, "N")
                'importe de venta / kg
                ImpVtaKg = 0
                If DBLet(RT!pesoneto, "N") <> 0 Then ImpVtaKg = Round2(DBLet(RT!impventa, "N") / DBLet(RT!pesoneto, "N"), 4)
                
                ExcelSheet.Cells(I, 31).Value = TransformaComasPuntos(ImpVtaKg)
                
                ' valor de la fruta
                ValorFruta = DBLet(RT!impventa, "N") - TotalGastos
                ExcelSheet.Cells(I, 32).Value = ValorFruta
                
                ' valor fruta / kg
                ImpVtaKg = 0
                If DBLet(RT!pesoneto, "N") <> 0 Then ImpVtaKg = Round2(ValorFruta / DBLet(RT!pesoneto, "N"), 4)
                ExcelSheet.Cells(I, 33).Value = TransformaComasPuntos(ImpVtaKg)
            
                ' facturado
                Select Case DBLet(RT!facturado, "N")
                    Case 0
                        ExcelSheet.Cells(I, 34).Value = "Provisional"
                    Case 1
                        ExcelSheet.Cells(I, 34).Value = "Definitivo"
                    Case 2
                        ExcelSheet.Cells(I, 34).Value = "Facturado"
                End Select
                
                ' cobrado
                If DBLet(RT!cobrado, "N") = 0 Then
                    ExcelSheet.Cells(I, 35).Value = "No"
                Else
                    ExcelSheet.Cells(I, 35).Value = "Sí"
                End If
    
                ' tipo de mercado
                Sql = "select nomtimer, tiptimer from tipomer where codtimer = " & DBSet(RsCAlb!codtimer, "N")
                Set RS = New ADODB.Recordset
                RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS.EOF Then
                    Select Case DBLet(RS.Fields(1).Value, "N")
                        Case 0
                            ExcelSheet.Cells(I, 36).Value = "Interior"
                        Case 1
                            ExcelSheet.Cells(I, 36).Value = "Exportación"
                        Case 2
                            ExcelSheet.Cells(I, 36).Value = "Industria"
                        Case 3
                            ExcelSheet.Cells(I, 36).Value = "Retirada"
                        Case 4
                            ExcelSheet.Cells(I, 36).Value = "Otros"
                    End Select
                    ExcelSheet.Cells(I, 37).Value = RS.Fields(1).Value
                            
                    ExcelSheet.Cells(I, 38).Value = RS.Fields(0).Value
                Else
                    ExcelSheet.Cells(I, 36).Value = ""
                    ExcelSheet.Cells(I, 37).Value = ""
                    ExcelSheet.Cells(I, 38).Value = ""
                End If
                Set RS = Nothing
    
            End If
        End If
        
        RT.MoveNext
    Wend
    
    RT.Close
    Set RT = Nothing
    
    RecorremosLineas = True
    
    Exit Function
    
eRecorremosLineas:
    Mens = Mens & vbCrLf & "Recorriendo lineas: " & Err.Description
End Function


Private Sub Image3_Click()
 AbrirDialogo 2
End Sub


Private Sub Image4_Click()
'    Set frmC = New frmCal
'    frmC.Fecha = Now
'    If Text4.Text <> "" Then
'        If IsDate(Text4.Text) Then frmC.Fecha = CDate(Text4.Text)
'    End If
'    frmC.Show vbModal
'    Set frmC = Nothing
End Sub

Private Sub Image5_Click()
    MsgBox "Formato importe:   SOLO el punto decimal: 1.49", vbExclamation
End Sub

'Private Sub Text4_LostFocus()
'    Text4.Text = Trim(Text4.Text)
'    If Text4.Text <> "" Then
'        If IsDate(Text4.Text) Then
'            Text4.Text = Format(Text4.Text, "dd/mm/yyyy")
'        Else
'            MsgBox "Fecha incorrecta", vbExclamation
'            Text4.Text = ""
'        End If
'    End If
'End Sub
'
'

'-------------------------------------
Private Function RecorremosLineasLiquidacion()
Dim FIN As Boolean
Dim I As Long
Dim JJ As Integer
Dim LineasEnBlanco As Integer

    'Desde la fila donde empieza los trabajadores
    'Hasta k este vacio
    'Iremos insertando en tmpHoras
    ' Con trbajador, importe, 0 , 1 ,2
    '             Existe, No existe, IMPORTE negativo
    '
    
    Sql = "DELETE FROM tmpExcel where codusu = " & Usuario
    Conn.Execute Sql
    FIN = False
    I = 2
    LineasEnBlanco = 0
    While Not FIN
        'Debug.Print "L: " & i
        If Trim(CStr(ExcelSheet.Cells(I, 1).Value)) <> "" Then
            LineasEnBlanco = 0
            If IsNumeric((ExcelSheet.Cells(I, 1).Value)) Then
                If Val(ExcelSheet.Cells(I, 1).Value) > 0 Then
                        'albaran
                        Albaran = Val(ExcelSheet.Cells(I, 1).Value)
                        
                        'Importe
                        FecAlbaran = Format(ExcelSheet.Cells(I, 2).Value, "yyyy/mm/dd")
                        Socio = ExcelSheet.Cells(I, 3).Value
                        Campo = ExcelSheet.Cells(I, 5).Value
                        Variedad = ExcelSheet.Cells(I, 6).Value
                        TipoEntr = ExcelSheet.Cells(I, 8).Value
                        KilosNet = ExcelSheet.Cells(I, 9).Value
                        
                        
                        For JJ = 1 To 20
                            Calidad(JJ) = Val(ExcelSheet.Cells(I, 9 + JJ).Value)
                        Next JJ
                        
                        'InsertartmpLiquida
                        InsertaTmpExcel
                    
                    End If
            End If
        Else
            LineasEnBlanco = LineasEnBlanco + 1
            If LineasEnBlanco < 30 Then
               ' FIN = False
            Else
                FIN = True
               
            End If
        End If
        'Siguiente
        'Siguiente
        I = I + 1
    Wend
End Function




Private Sub InsertaTmpExcel()
Dim vSQL As String
Dim vSql2 As String
Dim RT As ADODB.Recordset
Dim RT1 As ADODB.Recordset
Dim RT2 As ADODB.Recordset
Dim Existe As Boolean
Dim ExisteCalidad As Boolean
Dim ExisteEnTemporal As Boolean
Dim TotalKilos As Long
Dim Cuadra As Boolean
Dim JJ As Integer

    On Error GoTo EInsertaTmpExcel
    
    vSQL = "Select * from rhisfruta "
    vSQL = vSQL & " WHERE numalbar = " & Albaran
    vSQL = vSQL & " and fecalbar = '" & Format(CDate(FecAlbaran), "yyyy-mm-dd") & "'"
    vSQL = vSQL & " and codsocio = " & Socio
    vSQL = vSQL & " and codcampo = " & Campo
    vSQL = vSQL & " and codvarie = " & Variedad
    vSQL = vSQL & " and tipoentr = " & TipoEntr

    Set RT = New ADODB.Recordset
    RT.Open vSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If RT.EOF Then
        Existe = False
    Else
        Existe = True
    End If
    
    ' si existe la entrada vemos si podemos actualizarla
    If Existe Then
        ExisteCalidad = True
        
        For JJ = 1 To mConfig.MaxCalidades
            If Calidad(JJ) <> 0 Then  ' solo si hay kilos
'                vSQL = "select * from rhisfruta_clasif where numalbar = " & Albaran
'                vSQL = vSQL & " and codvarie = " & Variedad
'                vSQL = vSQL & " and codcalid = " & JJ
'
'                Set RT1 = New ADODB.Recordset
'                RT1.Open vSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                If RT1.EOF Then
                    vSql2 = "select * from rcalidad where codvarie = " & Variedad
                    vSql2 = vSql2 & " and codcalid = " & JJ
                    
                    Set RT2 = New ADODB.Recordset
                    RT2.Open vSql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
                    If RT2.EOF Then
                        ExisteCalidad = False
                        Set RT2 = Nothing
                        Exit For
                    Else
                        ExisteCalidad = True
                        Set RT2 = Nothing
                    End If
'                End If
                
            End If
        Next JJ
    
    
        If ExisteCalidad Then ' comprobamos que la suma de calidades da kilosnetos
            TotalKilos = 0
            For JJ = 1 To 20
                TotalKilos = TotalKilos + Calidad(JJ)
            Next JJ
            If TotalKilos <> RT!KilosNet Then
                Cuadra = False
            Else
                Cuadra = True
            End If
        End If
    
    End If
    
    If Existe And ExisteCalidad And Cuadra Then
        
        ExisteEnTemporal = False
        vSQL = "select * from tmpexcel where numalbar = " & Albaran & " and codusu = " & Usuario
    
        Set RT2 = New ADODB.Recordset
        RT2.Open vSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
        If Not RT2.EOF Then
            ExisteEnTemporal = True
        End If
    
        Sql = "insert into tmpexcel (codusu, numalbar, fecalbar, codsocio, codcampo, codvarie, tipoentr, kilosnet, "
        Sql = Sql & "calidad1, calidad2, calidad3, calidad4, calidad5, calidad6, calidad7, calidad8, calidad9, "
        Sql = Sql & "calidad10, calidad11, calidad12, calidad13, calidad14, calidad15, calidad16, calidad17, "
        Sql = Sql & "calidad18, calidad19, calidad20, situacion) values ("
        Sql = Sql & Usuario & ","
        Sql = Sql & Albaran & ","
        Sql = Sql & "'" & Format(CDate(FecAlbaran), "yyyy-mm-dd") & "',"
        Sql = Sql & Socio & ","
        Sql = Sql & Campo & ","
        Sql = Sql & Variedad & ","
        Sql = Sql & TipoEntr & ","
        Sql = Sql & KilosNet & ","
        
        For JJ = 1 To mConfig.MaxCalidades
            Sql = Sql & Calidad(JJ) & ","
        Next JJ
    
        If ExisteEnTemporal Then
            Sql = Sql & "2)"
        Else
            Sql = Sql & "0)"
        End If
        
    Else
        Sql = "insert into tmpexcel (codusu, numalbar, fecalbar, codsocio, codcampo, codvarie, tipoentr, kilosnet, "
        Sql = Sql & "calidad1, calidad2, calidad3, calidad4, calidad5, calidad6, calidad7, calidad8, calidad9, "
        Sql = Sql & "calidad10, calidad11, calidad12, calidad13, calidad14, calidad15, calidad16, calidad17,"
        Sql = Sql & "calidad18, calidad19, calidad20, situacion) values ("
        Sql = Sql & Usuario & ","
        Sql = Sql & Albaran & ",'"
        Sql = Sql & Format(CDate(FecAlbaran), "yyyy-mm-dd") & "',"
        Sql = Sql & Socio & ","
        Sql = Sql & Campo & ","
        Sql = Sql & Variedad & ","
        Sql = Sql & TipoEntr & ","
        Sql = Sql & KilosNet & ","
        
        For JJ = 1 To mConfig.MaxCalidades
            Sql = Sql & Calidad(JJ) & ","
        Next JJ
    
        If Not Existe Then
            Sql = Sql & "1)" ' no existe el albaran
        Else
            If Not ExisteCalidad Then ' no existe la calidad
                Sql = Sql & "11)"
            Else
                Sql = Sql & "12)" ' no cuadran kilos
            End If
        End If
        
    End If
    
    
    If Sql <> "" Then Conn.Execute Sql
        
    RT.Close
    
    Exit Sub
EInsertaTmpExcel:
    MsgBox Err.Description
End Sub



Private Sub vConfiguracion(Leer As Boolean)

'    With varConfig
'        If Leer Then
'            Text8(0).Text = .IniLinNomina
'            Text8(1).Text = .FinLinNominas
'            Text8(2).Text = .ColTrabajadorNom
'            Text8(3).Text = .hc
'            Text8(4).Text = .HPLUS
'            Text8(5).Text = .DIAST
'            Text8(6).Text = .Anticipos
'            Text8(7).Text = .ColTrabajadoresLIQ
'            Text8(8).Text = .ColumnaLiquidacion
'            Text8(9).Text = .FilaLIQ
'            Text8(10).Text = .HN
'        Else
'            .IniLinNomina = Val(Text8(0).Text)
'            .FinLinNominas = Val(Text8(1).Text)
'            .ColTrabajadorNom = Val(Text8(2).Text)
'            .hc = Val(Text8(3).Text)
'            .HPLUS = Val(Text8(4).Text)
'            .DIAST = Val(Text8(5).Text)
'            .Anticipos = Val(Text8(6).Text)
'            .ColTrabajadoresLIQ = Val(Text8(7).Text)
'            .ColumnaLiquidacion = Val(Text8(8).Text)
'            .FilaLIQ = Val(Text8(9).Text)
'            .HN = Val(Text8(10).Text)
'        End If
'    End With
End Sub

Private Sub Text8_GotFocus(Index As Integer)
    With Text8(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text8_LostFocus(Index As Integer)
    With Text8(Index)
        .Text = Trim(.Text)
        If .Text = "" Then Exit Sub
        
        Select Case Index
            Case 0 ' numero de calidades
                If Not IsNumeric(.Text) Then
                    MsgBox "Campo debe ser numérico", vbExclamation
                    .Text = ""
                    .SetFocus
                    Exit Sub
                End If
                .Text = Val(.Text)
            
            Case 2, 3 ' usuario y password deben de estar encriptados
            
            
        End Select
            
            
    End With
End Sub


Private Function Cooperativa() As Byte
Dim Sql As String
Dim RS As ADODB.Recordset

    On Error Resume Next

    Cooperativa = 0
    Sql = "select cooperativa from rparam "
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
    If Not RS.EOF Then
        Cooperativa = (DBLet(RS!Cooperativa, "N"))
    End If
    Set RS = Nothing

End Function

Private Function RecorremosLineasCatadau(Mens As String) As Boolean
Dim I As Integer
Dim j As Integer
Dim JJ As Integer
Dim F As Date
Dim Cod As String
Dim FE As String
Dim RT As ADODB.Recordset
Dim RS As ADODB.Recordset
Dim RsCAlb As ADODB.Recordset
Dim RsLAlb As ADODB.Recordset
Dim Calidad As Integer
Dim NFic As Integer
Dim Lin As String
Dim Sql1 As String
Dim Sql2 As String

Dim NFile As Integer
Dim TotalGastos As Currency
Dim GastosKg As Currency
Dim ImpVtaKg As Currency
Dim ValorFruta As Currency

    On Error GoTo eRecorremosLineas

    RecorremosLineasCatadau = False


    Sql = "select * from tmpinformes where codusu = " & Usuario & " order by importeb1, fecha1 "
    Sql1 = "select count(*) from tmpinformes where codusu = " & Usuario

    Set RT = New ADODB.Recordset
    RT.Open Sql1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RT.EOF Then
        Me.Pb1.visible = True
        Me.Pb1.Max = RT.Fields(0).Value
        Me.Pb1.Value = 0
        Me.Refresh
    End If
    
    Set RT = Nothing
    Set RT = New ADODB.Recordset
    RT.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    I = 1
    While Not RT.EOF
        I = I + 1
            
        IncrementarProgresNew Pb1, 1
    
        Sql2 = "select * from albaran where numalbar = " & DBSet(RT!importeb1, "N")
        Set RsCAlb = New ADODB.Recordset
        RsCAlb.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not RsCAlb.EOF Then
    
            Sql2 = "select * from albaran_variedad where numalbar = " & DBSet(RT!importeb1, "N") & " and numlinea = " & DBSet(RT!importeb2, "N")
            Set RsLAlb = New ADODB.Recordset
            RsLAlb.Open Sql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
            If Not RsLAlb.EOF Then
                
                ExcelSheet.Cells(I, 1).Value = DBLet(RsLAlb!numalbar, "N") 'nro de expediente
                ExcelSheet.Cells(I, 2).Value = DBLet(RsCAlb!fechaalb, "F") 'fecha
                ExcelSheet.Cells(I, 3).Value = DBLet(RsCAlb!matriveh, "T") 'matricula
                
                
                Sql = "select nomclien from clientes where codclien = " & DBSet(RsCAlb!codclien, "N")
                Set RS = New ADODB.Recordset
                RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS.EOF Then
                    ExcelSheet.Cells(I, 4).Value = RS.Fields(0).Value ' nombre de cliente
                Else
                    ExcelSheet.Cells(I, 4).Value = "" ' nombre de cliente
                End If
                Set RS = Nothing
                
                ' destino
                Sql = "select nomdesti from destinos where codclien = " & DBSet(RsCAlb!codclien, "N")
                Sql = Sql & " and coddesti = " & DBSet(RsCAlb!codDESTI, "N")
                Set RS = New ADODB.Recordset
                RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS.EOF Then
                    ExcelSheet.Cells(I, 5).Value = RS.Fields(0).Value ' nombre de destino
                Else
                    ExcelSheet.Cells(I, 5).Value = "" ' nombre de destino
                End If
                Set RS = Nothing
                
                ' calibre
                Sql = "select nomcalib from calibres where codvarie = " & DBSet(RsLAlb!codvarie, "N")
                Sql = Sql & " and codcalib = " & DBSet(RT!importe2, "N")
                Set RS = New ADODB.Recordset
                RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS.EOF Then
                    ExcelSheet.Cells(I, 6).Value = RS.Fields(0).Value ' nombre de calibre
                Else
                    ExcelSheet.Cells(I, 6).Value = "" ' nombre de calibre
                End If
                Set RS = Nothing
                
                ' envase
                Sql = "select nomconfe from forfaits where codforfait = " & DBSet(RsLAlb!codforfait, "T")
                Set RS = New ADODB.Recordset
                RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS.EOF Then
                    ExcelSheet.Cells(I, 7).Value = RS.Fields(0).Value ' nombre de forfait
                Else
                    ExcelSheet.Cells(I, 7).Value = "" ' nombre de forfait
                End If
                Set RS = Nothing
                
                'forfait
                ExcelSheet.Cells(I, 8).Value = ExcelSheet.Cells(I, 7).Value
                
                'marca
                Sql = "select nommarca from marcas where codmarca = " & DBSet(RsLAlb!codmarca, "N")
                Set RS = New ADODB.Recordset
                RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS.EOF Then
                    ExcelSheet.Cells(I, 9).Value = RS.Fields(0).Value ' nombre de marca
                Else
                    ExcelSheet.Cells(I, 9).Value = "" ' nombre de marca
                End If
                Set RS = Nothing
                
                'categoria
                ExcelSheet.Cells(I, 10).Value = DBLet(RsLAlb!categori, "N")  'categoria
                
                'tipo de palet
                Sql = "select nompalet from confpale where codpalet = " & DBSet(RsLAlb!codpalet, "N")
                Set RS = New ADODB.Recordset
                RS.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS.EOF Then
                    ExcelSheet.Cells(I, 11).Value = RS.Fields(0).Value ' tipo de palet
                Else
                    ExcelSheet.Cells(I, 11).Value = "" ' tipo de palet
                End If
                Set RS = Nothing
                
                ' el numero de palets como es por calibre lo dejo en blanco
                ExcelSheet.Cells(I, 12).Value = ""
                
                ' numero de cajas
                ExcelSheet.Cells(I, 13).Value = DBLet(RT!importe1, "N")
                
                ' kilosnetos
                ExcelSheet.Cells(I, 14).Value = DBLet(RT!importe3, "N")
                
                ' precio
                Dim Precio As Currency
                Precio = 0
                If DBLet(RT!importe3, "N") <> 0 Then
                    Precio = Round2(DBLet(RT!importe4, "N") / DBLet(RT!importe3, "N"), 4)
                End If
                ExcelSheet.Cells(I, 15).Value = Precio
                
            End If
        End If
        
        RT.MoveNext
    Wend
    
    RT.Close
    Set RT = Nothing
    
    RecorremosLineasCatadau = True
    
    Exit Function
    
eRecorremosLineas:
    Mens = Mens & vbCrLf & "Recorriendo lineas: " & Err.Description
End Function



