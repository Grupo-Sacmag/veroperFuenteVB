VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form VeroperI 
   Caption         =   "Ver operaciones"
   ClientHeight    =   8172
   ClientLeft      =   60
   ClientTop       =   288
   ClientWidth     =   13848
   LinkTopic       =   "Form1"
   ScaleHeight     =   8172
   ScaleWidth      =   13848
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame ArchTipo 
      Caption         =   "Tipo de archivo :"
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   5775
      Begin VB.OptionButton Option2 
         Caption         =   "Catalogos (Cataux,Catmay)"
         Height          =   255
         Left            =   3240
         TabIndex        =   5
         Top             =   360
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Operaciones(Polizas)"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   6360
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6240
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6840
      TabIndex        =   2
      Top             =   720
      Width           =   4215
   End
   Begin MSFlexGridLib.MSFlexGrid COPIA1 
      Height          =   6735
      Left            =   6840
      TabIndex        =   1
      Top             =   1080
      Width           =   6135
      _ExtentX        =   10816
      _ExtentY        =   11875
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   16777215
   End
   Begin MSFlexGridLib.MSFlexGrid VER1 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   6135
      _ExtentX        =   10816
      _ExtentY        =   11875
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedCols       =   0
      BackColorBkg    =   16777215
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   0
      Width           =   5055
   End
   Begin VB.Menu Arch 
      Caption         =   "&Archivo"
      Begin VB.Menu ArchAbr 
         Caption         =   "&Abrir"
         Begin VB.Menu ArchF 
            Caption         =   "&Fuente"
         End
         Begin VB.Menu ArchC 
            Caption         =   "&Copia"
         End
      End
      Begin VB.Menu ArchGua 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu ArcImp 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu ArchSal 
         Caption         =   "&Salida"
      End
   End
   Begin VB.Menu Edi 
      Caption         =   "&Edicion"
      Begin VB.Menu EdiCop 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu Editsep1 
         Caption         =   "-"
      End
      Begin VB.Menu EditPeg 
         Caption         =   "&Pegar"
         Shortcut        =   ^V
      End
      Begin VB.Menu EditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu EditIser 
         Caption         =   "&Insertar"
      End
      Begin VB.Menu EditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu EditElim 
         Caption         =   "&Eliminar"
      End
   End
   Begin VB.Menu Ca_Ax 
      Caption         =   "&Cambiar Auxiliares"
      Begin VB.Menu Ca_Or 
         Caption         =   "&Origen"
      End
      Begin VB.Menu Ca_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu Ca_Dest 
         Caption         =   "&Destino"
      End
   End
   Begin VB.Menu BorrSdos 
      Caption         =   "&Borrar Saldos"
   End
   Begin VB.Menu Bus 
      Caption         =   "B&uscar "
      Begin VB.Menu BusDep 
         Caption         =   "&Depositos"
      End
      Begin VB.Menu BuscSep1 
         Caption         =   "-"
      End
      Begin VB.Menu BuscDatos 
         Caption         =   "&Modificar Datos"
      End
   End
   Begin VB.Menu CRP 
      Caption         =   "&Correccion Polizas"
      Begin VB.Menu OpCorr 
         Caption         =   "&Corregir Operaciones"
      End
   End
End
Attribute VB_Name = "VeroperI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim difer As Long, temp As String, temp1 As String, L As Long, i_n1 As Long, z1
Dim valcelant, archivo1, importe As Currency, Pantalla As Integer, FondoCol As Long
Dim A As Integer, B As Integer, Colores As Long, Cta_CMay As Integer, Cta_Caux As Integer
Sub LocDep()
     Dim ImpCLTE As Currency, ImpBCO As Currency
     Dim AX As Integer, EnTre As Integer, YaEsTA As Integer
     Close 1
     Open "CATAUX" For Random As 1 Len = Len(CATAUX)
     CM = LOF(1) / Len(CATAUX)
     COPIA1.Clear
     COPIA1.Rows = 1
     For r = 1 To VER1.Rows - 1
            ImpBCO = 0: ImpCLTE = 0: YaEsTA = 0
            If VER1.TextMatrix(r, 5) = "A" Then
               r = r + 1
                Do Until (VER1.TextMatrix(r, 5) = "A") Or (r = VER1.Rows - 1)
                   
                  Select Case (VER1.TextMatrix(r, 1))
                  Case Is = 1101
                     ImpBCO = VER1.TextMatrix(r, 4)
                  Case Is = 1103
                    If (ImpBCO <> 0) Then
                     If YaEsTA = 0 Then
                      COPIA1.AddItem r & Chr(9) & "" & Chr(9) & "Deposito" & Chr(9) & " " & Chr(9) & _
                                   Format(ImpBCO, "###,###,##0.00")
                         YaEsTA = 1
                     End If
                            r = r + 1
                      Do Until VER1.TextMatrix(r, 5) <> "C"
                            EnTre = 1
                            ImpCLTE = VER1.TextMatrix(r, 4)
                            AX = VER1.TextMatrix(r, 1)
                            Get 1, AX, CATAUX
                            COPIA1.AddItem r & Chr(9) & CATAUX.CTA & Chr(9) & CATAUX.reda _
                                & Chr(9) & " " & Chr(9) & Format(ImpCLTE, "###,###,##0.00")
                            r = r + 1
                      Loop
                      If EnTre = 1 Then r = r - 1
                   Else:
                    Rem NADA
                  End If
                End Select
                  r = r + 1
                Loop
                If VER1.TextMatrix(r, 5) = "A" Then r = r - 1
            End If
     Next r
End Sub
Sub Guarda()
  Dim XX
  'Debug.Print DM; CM
  If Option1 = True Then
     Rem Nada
     Else
     Rem ESTA ES LA OPCION PARA LOS CATALOGOS OPTION2 = TRUE
     XX = MsgBox("Desea conservar el tamaño del archivo", vbYesNo, "GUARDANDO " + archivo1)
     If XX = vbYes Then
            COPIA1.Rows = VER1.Rows
            Else
            Rem SE RECORTA EL TAMAÑO DEL ARCHIVO
     End If
  End If
  
  If Option1 = True Then
  Rem ESTA ES OPCION PARA ARCHIVO DE OPERACIONES
  For r = 1 To COPIA1.Rows - 1
       If (COPIA1.TextMatrix(r, 1) > "") And (COPIA1.TextMatrix(r, 4) <> "") Then
            oper.CTA = COPIA1.TextMatrix(r, 1)
            oper.reda = Trim(COPIA1.TextMatrix(r, 2))
            oper.fecha = COPIA1.TextMatrix(r, 3)
            If COPIA1.TextMatrix(r, 4) <> "" Then
                importe = COPIA1.TextMatrix(r, 4)
                Else
                importe = 0
            End If
            oper.impo = Str(importe)
            oper.clav = COPIA1.TextMatrix(r, 5)
            oper.Real = COPIA1.TextMatrix(r, 6)
            Put 4, r, oper
       End If
  Next r
  Else
  For r = 1 To COPIA1.Rows - 1
       If (COPIA1.TextMatrix(r, 1) > "") And (COPIA1.TextMatrix(r, 4) <> "") Then
            otros.CTA = COPIA1.TextMatrix(r, 1)
            otros.reda = Trim(COPIA1.TextMatrix(r, 2))
            otros.fecha = COPIA1.TextMatrix(r, 3)
            If COPIA1.TextMatrix(r, 4) <> "" Then
                importe = COPIA1.TextMatrix(r, 4)
                Else
                importe = 0
            End If
            otros.impo = Str(importe)
            otros.clav = COPIA1.TextMatrix(r, 5)
            otros.Real = COPIA1.TextMatrix(r, 6)
            Put 4, r, otros
       
       Else
           Rem If COPIA1.Rows = (CM + 1) Then
                otros.CTA = ""
                otros.reda = ""
                otros.fecha = ""
                otros.impo = Str(0)
                otros.clav = ""
                otros.Real = ""
                Put 4, r, otros
                Rem Else
                Rem  NADA SE RECORTA EL ARCHIVO
       End If
      
  Next r
  End If
  Close 4
End Sub
Sub cargado()
      COPIA1.Clear
      COPIA1.Rows = 1
      If Option1 = True Then
            archivo1 = CommonDialog1.FileName
            Close 4: Open archivo1 For Random As 4 Len = Len(oper)
            DM = LOF(4) / Len(oper)
            COPIA1.Cols = 7
            COPIA1.Rows = 1
            COPIA1.ColWidth(1) = 800
            COPIA1.ColWidth(2) = 2300
            COPIA1.ColWidth(3) = 400
            COPIA1.ColWidth(4) = 1200
            COPIA1.ColWidth(5) = 200
            COPIA1.ColWidth(6) = 500
            COPIA1.ColWidth(0) = 800
            COPIA1.Width = 6500 + 800
        For r = 1 To DM
            Get 4, r, oper
            Real.CTA = Val(oper.CTA)
            Real.reda = oper.reda
            Real.fecha = oper.fecha
            Real.impo = Val(oper.impo)
            Real.clav = oper.clav
            Real.Real = Val(oper.Real)
            COPIA1.AddItem Format(r, "#####") & Chr(9) & Format(Real.CTA, "#####") & Chr(9) & " " + Real.reda _
                     & Chr(9) & Real.fecha & Chr(9) & Format(Real.impo, z1) _
                    & Chr(9) & Real.clav & Chr(9) & Format(Real.Real, "#####")
       
       If Real.clav = "A" Then
        COPIA1.BackColor = vbBlue
        Else
            COPIA1.BackColor = vbWhite
            COPIA1.ForeColor = vbBlack
        End If
 
        Next r
        COPIA1.Rows = COPIA1.Rows + VER1.Rows
      Else
      
      archivo1 = CommonDialog1.FileName
      Close 4: Open archivo1 For Random As 4 Len = Len(otros1)
      DM = LOF(4) / Len(otros1)

            COPIA1.Cols = 7
            COPIA1.Rows = 1
            COPIA1.ColWidth(1) = 800
            COPIA1.ColWidth(2) = 2300
            COPIA1.ColWidth(3) = 400
            COPIA1.ColWidth(4) = 1200
            COPIA1.ColWidth(5) = 500
            COPIA1.ColWidth(6) = 500
            COPIA1.ColWidth(0) = 800
            COPIA1.Width = 6500 + 800
        For r = 1 To DM
            Get 4, r, otros1
            Real1.CTA = Val(otros1.CTA)
            Real1.reda = otros1.reda
            Real1.fecha = otros1.fecha
            Real1.impo = Val(otros1.impo)
            Real1.clav = otros1.clav
            Real1.Real = Val(otros1.Real)
            COPIA1.AddItem Format(r, "#####") & Chr(9) & Format(Real1.CTA, "#####") & Chr(9) & " " + Real1.reda _
                     & Chr(9) & Real1.fecha & Chr(9) & Format(Real1.impo, z1) _
                    & Chr(9) & Real1.clav & Chr(9) & Format(Real1.Real, "#####")
       
       If Real1.clav = "A" Then
        
        COPIA1.BackColor = vbBlue
        
        Else
            COPIA1.BackColor = vbWhite
            COPIA1.ForeColor = vbBlack
            End If
 
        Next r
        COPIA1.Rows = COPIA1.Rows + VER1.Rows
  End If
End Sub
Sub CARGA()
  On Error GoTo MANEJOERR
   VER1.Clear
   VER1.Rows = 1
   VER1.Cols = 7
   VER1.ColWidth(1) = 800
   VER1.ColWidth(2) = 2300
   VER1.ColWidth(3) = 400
   VER1.ColWidth(4) = 1200
   If Option1.Value = True Then
        VER1.ColWidth(5) = 200
        VER1.ColWidth(6) = 500
        Else
        VER1.ColWidth(5) = 500
        VER1.ColWidth(6) = 500
   End If
   VER1.ColWidth(0) = 800
   VER1.Width = 6500 + 800
   COPIA1.Cols = 7
   COPIA1.ColWidth(1) = 800
   COPIA1.ColWidth(2) = 2300
   COPIA1.ColWidth(3) = 400
   COPIA1.ColWidth(4) = 1200
   If Option1.Value = 1 Then
        COPIA1.ColWidth(5) = 200
        COPIA1.ColWidth(6) = 500
        Else
        COPIA1.ColWidth(5) = 500
        COPIA1.ColWidth(6) = 500
   End If
   COPIA1.ColWidth(0) = 800
   archivo1 = CommonDialog1.FileName
   If Option1 = True Then
        VeroperI.Caption = "VER OPERACIONES ARCHIVO: " + archivo1
        Close 1: Open archivo1 For Random As 1 Len = Len(oper)
        CM = LOF(1) / Len(oper)
        For r = 1 To CM
        Get 1, r, oper
        Real.CTA = Val(oper.CTA): Rem 1
        Real.reda = oper.reda: Rem 2
        Real.fecha = oper.fecha: Rem 3
        Real.impo = Val(oper.impo): Rem 4
        Real.clav = oper.clav: Rem 5
        Real.Real = Val(oper.Real): Rem 6
        VER1.AddItem Format(r, "#####") & Chr(9) & Format(Real.CTA, "#####") & Chr(9) & " " + Real.reda _
            & Chr(9) & Real.fecha & Chr(9) & Format(Real.impo, z1) _
            & Chr(9) & Real.clav & Chr(9) & Format(Real.Real, "#####")
       
       If Real.clav = "A" Then
        
        VER1.CellBackColor = vbBlue
        
        Else
         VER1.CellBackColor = vbWhite
         VER1.ForeColor = vbBlack
       End If
 
   Next r
   Else
        
        VeroperI.Caption = "VER CATALOGO ARCHIVO: " + archivo1
        Close 1: Open archivo1 For Random As 1 Len = Len(otros)
        CM = LOF(1) / Len(otros)
        For r = 1 To CM
        Get 1, r, otros
        Real1.CTA = Val(otros.CTA)
        Real1.reda = otros.reda
        Real1.fecha = otros.fecha
        Real1.impo = Val(otros.impo)
        Real1.clav = otros.clav
        Real1.Real = Val(otros.Real)
        VER1.AddItem Format(r, "#####") & Chr(9) & Format(Real1.CTA, "#####") & Chr(9) & " " + Real1.reda _
            & Chr(9) & Real1.fecha & Chr(9) & Format(Real1.impo, z1) _
            & Chr(9) & Real1.clav & Chr(9) & Format(Real1.Real, "#####")
       
       If Real1.clav = "A" Then
        
        VER1.CellBackColor = vbBlue
        
        Else
         VER1.CellBackColor = vbWhite
         VER1.ForeColor = vbBlack
       End If
 
   Next r
   End If
   COPIA1.Rows = VER1.Rows
MANEJOERR:
Close
End Sub
Sub CHECA()
 Dim ClpFmt, Msg   ' Declara variables.
   On Error Resume Next   ' Configura el tratamiento de errores.
   If Clipboard.GetFormat(vbCFText) Then ClpFmt = ClpFmt + 1
   If Clipboard.GetFormat(vbCFBitmap) Then ClpFmt = ClpFmt + 2
   If Clipboard.GetFormat(vbCFDIB) Then ClpFmt = ClpFmt + 4
   If Clipboard.GetFormat(vbCFRTF) Then ClpFmt = ClpFmt + 8
   Select Case ClpFmt
      Case 1
         Msg = "El Portapapeles sólo contiene texto."
      Case 2, 4, 6
         Msg = "El Portapapeles sólo contiene un mapa de bits."
      Case 3, 5, 7
         Msg = "El Portapapeles contiene texto y un mapa de bits."
      Case 8, 9
         Msg = "El Portapapeles sólo contiene texto enriquecido."
      Case Else
         Msg = "No hay nada en el Portapapeles."
   End Select
   MsgBox Msg   ' Muestra el mensaje
End Sub

Private Sub ArchC_Click()
   If (Option1 = False) And (Option2 = False) Then
      MsgBox "Tiene que elegir el tipo de archivo que se va a usar", vbCritical
   Else
   On Err GoTo ErrHand
            CommonDialog1.Flags = cdlOFNHideReadOnly
            If Option1 = True Then
                CommonDialog1.FileName = "*.*"
                CommonDialog1.Filter = "Archivos de Operaciones(*.*)|*.*"
                Else
                CommonDialog1.FileName = "CAT*.*"
                CommonDialog1.Filter = "Archivos de catalogos(CAT*.*)|*.*"
            End If
            CommonDialog1.ShowOpen
            If CommonDialog1.FileName <> "" Then
               cargado
                    Else
                    nombrearchivo = ""
            End If
    End If
ErrHand:
   Exit Sub

End Sub

Private Sub ArchF_Click()
  If (Option1 = False) And (Option2 = False) Then
      MsgBox "Tiene que elegir el tipo de archivo que se va a usar", vbCritical
  Else
  On Err GoTo ErrHandler
            CommonDialog1.Flags = cdlOFNHideReadOnly
            If Option1 = True Then
                    CommonDialog1.FileName = "*.*"
                    CommonDialog1.Filter = "Archivos de Operaciones(*.*)|*.*"
                    Else
                    CommonDialog1.FileName = "CAT*.*"
                    CommonDialog1.Filter = "Archivos de catalogos(CAT*.*)|*.*"
            End If
            CommonDialog1.ShowOpen
            If CommonDialog1.FileName <> "" Then
               CARGA
                    Else
                    nombrearchivo = ""
            End If
  End If
ErrHandler:
   Close
   Exit Sub

End Sub

Private Sub ArchGua_Click()
  On Err GoTo manejo
  
         CommonDialog1.Filter = "Archivos de Operaciones(*.*)|*.*"
         CommonDialog1.ShowSave
            If CommonDialog1.FileName <> "" Then
               Close 4
               
               If Option1 = True Then
                    Open CommonDialog1.FileName For Random As 4 Len = Len(oper)
                    Close: Kill CommonDialog1.FileName
                    Open CommonDialog1.FileName For Random As 4 Len = Len(oper)
                    Else
                    Open CommonDialog1.FileName For Random As 4 Len = Len(otros)
                    Close: Kill CommonDialog1.FileName
                    Open CommonDialog1.FileName For Random As 4 Len = Len(otros)
               End If
               Guarda
                    Else
                    nombrearchivo = ""
            End If
manejo:
   Exit Sub
   
End Sub

Private Sub ArcImp_Click()
  For I = 1 To COPIA1.RowSel
     For h = 0 To 4
      Printer.Print COPIA1.TextMatrix(I, h);
     Next h
     Printer.Print
     
  Next I
  If COPIA1.RowSel > 1 Then Printer.EndDoc
End Sub

Private Sub ArchSal_Click()
 Close: End
End Sub

Private Sub BorrSdos_Click()
Dim Mensaje, Estilo, Título, Ayuda, Ctxt, Respuesta, MiCadena
Mensaje = "Desea borrar los saldos de este archivo ???"   ' Define el mensaje.
Estilo = vbYesNo + vbCritical + vbDefaultButton2   ' Define los botones.
Título = "Borrar "   ' Define el título.
   ' Define el archivo de ayuda.
   ' Define el tema
            ' el contexto
  If Option2 = True Then
      Respuesta = MsgBox(Mensaje, Estilo, Título)
      If Respuesta = vbYes Then   ' El usuario eligió el botón Sí.
    ' Ejecuta una acción.
      Close 4: Open archivo1 For Random As 4 Len = Len(otros1)
      DM = LOF(4) / Len(otros1)
            COPIA1.Cols = 7
            COPIA1.Rows = 1
            COPIA1.ColWidth(1) = 800
            COPIA1.ColWidth(2) = 2300
            COPIA1.ColWidth(3) = 400
            COPIA1.ColWidth(4) = 1200
            COPIA1.ColWidth(5) = 500
            COPIA1.ColWidth(6) = 500
            COPIA1.ColWidth(0) = 800
            COPIA1.Width = 6500 + 800
        For r = 1 To DM
            Get 4, r, otros1
            Real1.CTA = Val(otros1.CTA)
            Real1.reda = otros1.reda
            Real1.fecha = otros1.fecha
            Real1.impo = Val(0)
            Real1.clav = otros1.clav
            Real1.Real = Val(otros1.Real)
            COPIA1.AddItem Format(r, "#####") & Chr(9) & Format(Real1.CTA, "#####") & Chr(9) & " " + Real1.reda _
                     & Chr(9) & Real1.fecha & Chr(9) & Format(Real1.impo, z1) _
                    & Chr(9) & Real1.clav & Chr(9) & Format(Real1.Real, "#####")
       
       If Real1.clav = "A" Then
        
        COPIA1.BackColor = vbBlue
        
        Else
            COPIA1.BackColor = vbWhite
            COPIA1.ForeColor = vbBlack
            End If
 
        Next r
        COPIA1.Rows = COPIA1.Rows + VER1.Rows
        
     Else   ' El usuario eligió el botón No.
      MsgBox "Cancelando el borrado ", vbCritical
      Exit Sub
    End If
    Else            ' Muestra el mensaje.
    MsgBox "Tiene que elegir catalogo de ctas para esta opcion", vbCritical
      Exit Sub   ' Ejecuta una acción.
  End If
    
End Sub

Private Sub BuscDatos_Click()
   VDatos.Show 1
End Sub

Private Sub BusDep_Click()
   Close
   LocDep
End Sub

Private Sub Ca_Dest_Click()
   If A <= 0 Then
     MsgBox "No se eligio cuenta de origen"
     Exit Sub
     Else
     B = InputBox("Cambiando cuentas ", "Numero de cuenta de destino ")
       Open "CATMAY" For Random As 1 Len = Len(otros)
       CM = LOF(1) / Len(otros)
       For r = 1 To CM: Get 1, r, otros
                 If B = Val(otros.CTA) Then
                        Label1.BackColor = vbCyan
                        Label1.Caption = "Cambiando La Cuenta " + " " + Str(A) + _
                                         " A " + Str(B)
                        Dest.DE = Val(otros.clav)
                        Dest.A = Val(otros.Real)
                        Dest.Ubic = r
                        If (Dest.DE <= 0) Or (Dest.A < Dest.DE) Then
                           MsgBox "La cuenta no es correcta ", vbCritical
                           Label1.Caption = "": Label1.BackColor = Colores
                        End If

                        Exit For
                 End If
       Next r
       Close
       If A > 0 Then
          Cambio_Datos
          Cta_CMay = 0: Cta_Caux = 0
          Cambio_oper
          MsgBox "Se efectuaron " + Str(Cta_CMay) + " En ctas de Mayor " & Chr(13) & _
                 " y " + Str(Cta_Caux) + " en auxiliares ", vbCritical
          Label1.BackColor = Colores: A = 0: B = 0
          Dest.DE = 0: Dest.A = 0: Dest.Ubic = 0
          Orgs.A = 0: Orgs.DE = 0: Orgs.Ubic = 0
          Label1.Caption = ""
       End If
   End If
End Sub
Sub Cambio_Datos()
   Dim Resp, DiF_Cia As Integer
    Resp = MsgBox("DESEA TRASLADAR LOS NOMBRES DE LOS " & Chr(13) & _
           " AUXILIARES A LA NUEVA CUENTA ?", vbYesNo)
    If Resp = vbYes Then
        DiF_Cia = (Dest.DE - Orgs.DE)
        Open "CATAUX" For Random As 1 Len = Len(otros)
        CM = LOF(1) / Len(otros)
        For r = Orgs.DE To Orgs.A: Get 1, r, otros
            If IsNumeric(otros.CTA) <> 0 Then
                 Put 1, (r + DiF_Cia), otros
                 otros.clav = "": otros.CTA = "": otros.fecha = ""
                 otros.impo = "": otros.Real = "": otros.reda = ""
                 Put 1, r, otros
            End If
        Next r
        Close
        Else
         Rem NADA ****************
        Exit Sub
    End If
End Sub
Sub Cambio_oper()
    DiF_Cia = Dest.DE - Orgs.DE
    For r = 1 To VER1.Rows - 1
       Select Case (VER1.TextMatrix(r, 5))
            Case "A"
             Copia_Oper
            Case "B"
             If A = VER1.TextMatrix(r, 1) Then
                COPIA1.TextMatrix(r, 0) = VER1.TextMatrix(r, 0)
                COPIA1.TextMatrix(r, 1) = B
                COPIA1.TextMatrix(r, 2) = VER1.TextMatrix(r, 2)
                COPIA1.TextMatrix(r, 3) = VER1.TextMatrix(r, 3)
                COPIA1.TextMatrix(r, 4) = VER1.TextMatrix(r, 4)
                COPIA1.TextMatrix(r, 5) = VER1.TextMatrix(r, 5)
                COPIA1.TextMatrix(r, 6) = Dest.Ubic
                Cta_CMay = Cta_CMay + 1
                Else
                Copia_Oper
             End If
            Case "C"
              If (VER1.TextMatrix(r, 1) >= Orgs.DE) And (VER1.TextMatrix(r, 1) <= Orgs.A) Then
                COPIA1.TextMatrix(r, 0) = VER1.TextMatrix(r, 0)
                COPIA1.TextMatrix(r, 1) = (DiF_Cia + (VER1.TextMatrix(r, 1)))
                COPIA1.TextMatrix(r, 2) = VER1.TextMatrix(r, 2)
                COPIA1.TextMatrix(r, 3) = VER1.TextMatrix(r, 3)
                COPIA1.TextMatrix(r, 4) = VER1.TextMatrix(r, 4)
                COPIA1.TextMatrix(r, 5) = VER1.TextMatrix(r, 5)
                COPIA1.TextMatrix(r, 6) = VER1.TextMatrix(r, 6)
                Cta_Caux = Cta_Caux + 1
                Else
                Copia_Oper
              End If
            Case Else
             Copia_Oper
       End Select
       
       
    Next r
End Sub
Sub Copia_Oper()
    For I = 0 To 6
        COPIA1.TextMatrix(r, I) = VER1.TextMatrix(r, I)
    Next I
End Sub
Private Sub Ca_Or_Click()
    
    If (VER1.Rows < 2 Or Option2 = True) Then
       MsgBox "Elija Archivo de operaciones"
       Exit Sub
       Else
       A = InputBox("Cambiando cuentas ", "Numero de cuenta de origen ")
       Open "CATMAY" For Random As 1 Len = Len(otros)
       CM = LOF(1) / Len(otros)
       For r = 1 To CM: Get 1, r, otros
                 If A = Val(otros.CTA) Then
                        Colores = Label1.BackColor
                        Label1.BackColor = vbCyan
                        Label1.Caption = "Cambiando La Cuenta " + Str(A)
                        Orgs.DE = Val(otros.clav)
                        Orgs.A = Val(otros.Real)
                        Orgs.Ubic = r
                        If (Orgs.DE <= 0) Or (Orgs.A < Orgs.DE) Then
                           MsgBox "La cuenta no es correcta ", vbCritical
                           Label1.Caption = "": Label1.BackColor = Colores
                        End If
                        Exit For
                 End If
       Next r
       Close
    End If
End Sub

Private Sub OpCorr_Click()
     If (Option1 = False) And (Option2 = False) Then
      MsgBox "Tiene que elegir el tipo de archivo que se va a usar", vbCritical
  Else
  On Err GoTo ErrHandler1
     If VER1.Rows > 1 Then
            CorrPol.Show 1
            Else
            MsgBox "No hay archivo de operaciones  ", vbCritical
     End If
  End If
ErrHandler1:
   Close
      
   Exit Sub

    
End Sub

Private Sub option1_Click()
    Option2 = False
    Option2.BackColor = FondoCol
    Option1.BackColor = vbYellow
    Option1 = True
    VER1.Clear
    Text1.Text = ""
    Label1.Refresh
    COPIA1.Clear
    VER1.Rows = 1
    COPIA1.Rows = 1
End Sub

Private Sub option2_Click()
    Option1.BackColor = FondoCol
    Option1 = False
    Option2.BackColor = vbYellow
    Option2 = True
    VER1.Clear
    COPIA1.Clear
    Text1.Text = ""
    VER1.Rows = 1
    COPIA1.Rows = 1

End Sub

Private Sub COPIA1_GotFocus()
   Pantalla = 2
End Sub

Private Sub EdiCop_Click()
   Dim Temporal1
        
 If Pantalla = 1 Then
   Clipboard.Clear
   
   difer = VER1.RowSel - VER1.Row
   For I = VER1.Row To VER1.RowSel
      Rem For f = 0 To VER1.ColSel
      For f = VER1.Col To VER1.ColSel
            Temporal1 = Temporal1 + VER1.TextMatrix(I, f)
            If f < VER1.ColSel Then
                Temporal1 = Temporal1 & Chr(9)
            End If
      Next f
      Temporal1 = Temporal1 & Chr(13) & Chr(10)
      
   Next I
    Clipboard.SetText Temporal1
    
   
   Else
   Clipboard.Clear
   
   difer = COPIA1.RowSel - COPIA1.Row
   For I = COPIA1.Row To COPIA1.RowSel
      Rem For f = 0 To COPIA1.ColSel
      For f = COPIA1.Col To COPIA1.ColSel
            Temporal1 = Temporal1 + COPIA1.TextMatrix(I, f)
            If f < COPIA1.ColSel Then
                Temporal1 = Temporal1 & Chr(9)
            End If
      Next f
      Temporal1 = Temporal1 & Chr(13) & Chr(10)
      
   Next I
    Clipboard.SetText Temporal1
    
   End If
 

sale1:
End Sub

Private Sub EditElim_Click()
   
   Rep_Num = COPIA1.TextMatrix(COPIA1.Row, 0) - 1
   For w = COPIA1.Row To COPIA1.RowSel
            COPIA1.RemoveItem COPIA1.Row
   Next w
   CorrNum
   
End Sub

Private Sub EditIser_Click()
    For w = COPIA1.Row To COPIA1.RowSel
     COPIA1.AddItem (""), COPIA1.Row
    Next w
     CorrNum
End Sub

Private Sub EditPeg_Click()
 Rem CHECA
 i_n1 = 1
 Dim temporal, TexEdicion, deAqui As Long, RetornoCarro As Long, InicioCopia As Long
  temporal = Clipboard.GetText(vbCFText)
  RetornoCarro = COPIA1.Col
  InicioCopia = COPIA1.Row
If temporal <> "" Then
  Clipboard.Clear
  deAqui = 1
For I = 1 To Len(temporal)
    Select Case Mid(temporal, I, 1)
          Case Chr(9)
          TexEdicion = Mid(temporal, deAqui, (I - deAqui))
          COPIA1.Text = Mid(temporal, deAqui, (I - deAqui))
          COPIA1.Col = COPIA1.Col + 1
          deAqui = I + 1
          Case Chr(13)
          TexEdicion = Mid(temporal, deAqui, (I - deAqui))
          COPIA1.Text = Mid(temporal, deAqui, (I - deAqui))
          If (COPIA1.Rows - 1) > COPIA1.Row Then COPIA1.Row = COPIA1.Row + 1
          deAqui = I + 1
          Case Chr(10)
          COPIA1.Col = RetornoCarro
          deAqui = I + 1
          Case Else
          Rem nada
    End Select
 Next I
 COPIA1.Row = InicioCopia: COPIA1.Col = RetornoCarro
 CorrNum
End If
Clipboard.Clear
End Sub
Sub CorrNum()
 Dim Rep_Num As Long
   For w = 1 To COPIA1.Rows - 1
           Rep_Num = Rep_Num + 1
           COPIA1.TextMatrix(w, 0) = Rep_Num
   Next w
   
End Sub
Private Sub Form_Load()
   Mm(0) = "Incorporacion": Mm(1) = "Enero": Mm(2) = "Febrero": Mm(3) = "Marzo": Mm(4) = "Abril"
   Mm(5) = "Mayo": Mm(6) = "Junio": Mm(7) = "Julio": Mm(8) = "Agosto"
   Mm(9) = "Septiembre": Mm(10) = "Octubre": Mm(11) = "Noviembre": Mm(12) = "Diciembre"
   Mm(13) = "Incorporacion"
   FondoCol = Option1.BackColor
   z1 = "##,###,##0.00"
   VER1.Cols = 5
   VER1.ColWidth(0) = 800
   VER1.ColWidth(1) = 2700
   VER1.ColWidth(2) = 1200
   VER1.ColWidth(3) = 200
   VER1.ColWidth(4) = 1200
   VER1.Width = 6500
   COPIA1.Cols = 5
   COPIA1.ColWidth(0) = 800
   COPIA1.ColWidth(1) = 2700
   COPIA1.ColWidth(2) = 1200
   COPIA1.ColWidth(3) = 200
   COPIA1.ColWidth(4) = 1200
   COPIA1.Width = 6500
   Rem archivo1 = "\SOC2001\SOC10"
   VeroperI.Caption = VeroperI.Caption + " " + archivo1
   Rem Open archivo1 For Random As 1 Len = Len(oper)
   Rem cm = LOF(1) / Len(oper)
   Rem Open "\CORS2001\CORSOC08" For Random As 2 Len = Len(oper1)
   Rem dm = LOF(2) / Len(oper1)
   
   Rem For R = 1 To cm
       Rem Get 1, R, oper
       Rem real.cta = Val(oper.cta)
       Rem real.reda = oper.reda
       Rem real.fecha = oper.fecha
       Rem real.impo = Val(oper.impo)
       Rem real.clav = oper.clav
       Rem real.real = Val(oper.real)
       Rem VER1.AddItem Format(real.cta, "#####") & Chr(9) & real.reda _
            & Chr(9) & real.fecha & Chr(9) & Format(real.impo, z1) _
            & Chr(9) & real.clav & Chr(9) & Format(real.real, "#####")
       
       Rem If real.clav = "A" Then
        
        Rem VER1.CellBackColor = vbBlue
        
        Rem Else
         Rem VER1.CellBackColor = vbWhite
         Rem VER1.ForeColor = vbBlack
       Rem End If
 
   Rem Next R
   COPIA1.Rows = VER1.Rows
End Sub
Private Sub copia1_EnterCell()
   If COPIA1.Col > 0 And COPIA1.Row > 0 Then
       COPIA1.CellBackColor = vbCyan
       Text1.Text = COPIA1.Text
       
       
   End If
End Sub
Private Sub copia1_LeaveCell()
    If COPIA1.Col > 0 And COPIA1.Row > 0 Then
        COPIA1.CellBackColor = vbWhite
        
    End If

  End Sub

Private Sub copia1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
COPIA1.RowSel = COPIA1.Row
End Sub

Private Sub copia1_RowColChange()
     
     If COPIA1.Text <> "" Then valcelant = COPIA1.Text
     
End Sub

Private Sub copia1_KeyDown(KeyCode As Integer, Shift As Integer)
      Select Case KeyCode
            Case vbKeyDelete
              'respuesta = MsgBox("DESEA ELIMINAR ESTE RENGLON ", vbYesNo + vbCritical + vbDefaultButton2)
              'If respuesta = vbYes Then
                'copia1.RemoveItem copia1.Row
                Rem copia1.Text = ""
                
              'End If
                   
    
            If (COPIA1.RowSel > COPIA1.Row) Or (COPIA1.ColSel > COPIA1.Col) Then
              Text1.Text = ""
              COPIA1.Text = ""
              borrar_seleccion
            End If
                Text1.Text = ""
                COPIA1.Text = ""
                
            Case vbKeyF2
               
                If COPIA1.Text <> "" Then valcelant = COPIA1.Text
                
                Text1.Text = COPIA1.Text
                Text1.SetFocus
               
       End Select
End Sub

Private Sub copia1_KeyPress(KeyAscii As Integer)
        COPIA1.SelectionMode = flexSelectionFree
        
    
        If COPIA1.Text <> "" Then valcelant = COPIA1.Text
           Text1.Text = Chr(KeyAscii)
           
           Text1.SetFocus
End Sub
Sub borrar_seleccion()
   For L = COPIA1.Row To COPIA1.RowSel
         For C = COPIA1.Col To COPIA1.ColSel
             COPIA1.TextMatrix(L, C) = Text1.Text
         Next C
    Next L

End Sub


Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
     
    Select Case KeyAscii
       Case 27
        Text1.Text = valcelant
        COPIA1.Text = valcelant
        COPIA1.SetFocus

       Case 13
        Select Case (COPIA1.Col)
          Case 1, 2, 3
            COPIA1.Text = Text1.Text
            COPIA1.SetFocus
          Case 4
                     
            COPIA1.Text = Format(Text1.Text, z1)
            COPIA1.SetFocus
         Case 5
           If Option1 = True Then
            MsgBox "Esta columna no se puede modificar "
            Text1.Text = valcelant
            COPIA1.Text = valcelant
            Else
            COPIA1.Text = Text1.Text
            COPIA1.SetFocus
            End If
        End Select
       Case Else
        COPIA1.Text = Text1.Text
 
    End Select
  End Sub
Private Sub Text1_Change()
     If Text1.Text = Chr(27) Then Text1.Text = valcelant
     
     COPIA1.Text = Text1.Text
     
End Sub
Private Sub Text1_GotFocus()
    SendKeys "{end}"
    
End Sub

Private Sub VER1_GotFocus()
    Pantalla = 1
End Sub

Private Sub VER1_LostFocus()
  If VER1.Rows > 1 Then
    VER1.Row = VER1.Row: VER1.Col = VER1.Col
    VER1.CellBackColor = vbWhite
  End If
End Sub

