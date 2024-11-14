VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Aplper 
   Caption         =   "Aplicación proyectos"
   ClientHeight    =   6390
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   -480
      TabIndex        =   1
      Top             =   0
      Width           =   3855
   End
   Begin MSFlexGridLib.MSFlexGrid CapOb1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   9340
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      BackColorBkg    =   16777215
      Appearance      =   0
   End
   Begin VB.Menu Arch 
      Caption         =   "&Archivo"
      Begin VB.Menu ArcGua 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu aplsep1 
         Caption         =   "-"
      End
      Begin VB.Menu ArcSal 
         Caption         =   "&Salida"
      End
   End
   Begin VB.Menu Edi 
      Caption         =   "&Editar"
      Begin VB.Menu EdiBorr 
         Caption         =   "&Borrar"
         Shortcut        =   ^B
      End
      Begin VB.Menu EdiBto 
         Caption         =   "Borrar &Todo"
         Shortcut        =   ^T
      End
   End
End
Attribute VB_Name = "Aplper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim valcelant, selocalizo As Integer
Sub vertsum()
   Dim cantid As Currency
   For i = 2 To 3
     For r = 1 To 20
      If CapOb1.TextMatrix(r, i) <> "" Then
         cantid = cantid + CapOb1.TextMatrix(r, i)
      End If
     Next r
     CapOb1.TextMatrix(CapOb1.Rows - 1, i) = Format(cantid, z1)
     cantid = 0
   Next i
End Sub

Sub cargamaestro()
   For i = 1 To CapOb1.Rows - 1
      If CapOb1.TextMatrix(i, 1) <> "" Then
         obra(i) = CapOb1.TextMatrix(i, 1)
         porcentaje(i) = CapOb1.TextMatrix(i, 2)
         Else
         obra(i) = 0
         porcentaje(i) = 0
      
      End If
  Next i
End Sub


Private Sub ArcGua_Click()
   If CapOb1.TextMatrix(CapOb1.Rows - 1, 2) = 100 Then
      cargamaestro
      grabamaestro
      Put 8, tar, maestro
      ultimo.num = 1
      Else
      MsgBox "No se puede archivar" & Chr(10) & "El porcentaje execede el 100%"
   End If
End Sub

Private Sub ArcSal_Click()
   Unload Aplper
End Sub

Private Sub CapOb1_EnterCell()
 If (CapOb1.Row > 0) And (CapOb1.Col > 0) Then
  CapOb1.CellBackColor = vbYellow
 End If
   valcelant = CapOb1.Text
End Sub

Private Sub CapOb1_GotFocus()
    
    CapOb1.Col = CapOb1.Col
    CapOb1_EnterCell
End Sub

Private Sub CapOb1_KeyDown(KeyCode As Integer, Shift As Integer)
   
   Select Case KeyCode
    Case vbKeyDelete
        CapOb1.Text = ""
        Text1.Text = CapOb1.Text
    Case vbKeyF2
        Text1.Text = CapOb1.Text
        Text1.SetFocus
   End Select
End Sub

Private Sub CapOb1_KeyPress(KeyAscii As Integer)
   verentrada
   valcelant = CapOb1.Text
   Text1.Text = Chr(KeyAscii)
   Text1.SetFocus
End Sub

Private Sub CapOb1_LeaveCell()
If CapOb1.Rows > 1 Then
 If (CapOb1.Row > 0) And (CapOb1.Col > 0) Then
  CapOb1.CellBackColor = vbWhite
 End If
End If
End Sub

Sub verentrada()
   If CapOb1.Col = 3 Then
       CapOb1.Col = 2
       CapOb1_LeaveCell
       CapOb1_EnterCell
   End If
End Sub

Private Sub CapOb1_RowColChange()
    Text1.Text = CapOb1.Text
End Sub

Private Sub EdiBorr_Click()
    For i = 1 To 3
        CapOb1.TextMatrix(CapOb1.Row, i) = ""
    Next i
    vertsum
    Text1.Text = ""
End Sub

Private Sub EdiBto_Click()
  For f = 1 To CapOb1.Rows - 1
    For i = 1 To 3
        CapOb1.TextMatrix(f, i) = ""
    Next i
  Next f
  Text1.Text = ""
End Sub

Private Sub Form_Load()
     Aplper.Caption = Poliza.Apl1.TextMatrix(Poliza.Apl1.Row, 1)
     convierte
     Form_Resize
     Col_Def
     CapOb1.Row = 0
     For i = 1 To 20
       If obra(i) > 0 Then
            importe = porcentaje(i) * base / 100
            CapOb1.AddItem i & Chr(9) & obra(i) & Chr(9) & _
            porcentaje(i) & Chr(9) & Format(importe, z1)
            Else
            CapOb1.AddItem i & Chr(9) & "" & Chr(9) & _
            "" & Chr(9) & ""
       End If
     Next i
     CapOb1.Rows = CapOb1.Rows + 1
     CapOb1.TextMatrix(CapOb1.Rows - 1, 0) = "Totales"
     vertsum
     CapOb1.Row = 1: CapOb1.Col = 1:
     CapOb1_LeaveCell
     CapOb1_EnterCell
End Sub

Sub Col_Def()
  CapOb1.Row = 0
  CapOb1.Col = 0: CapOb1.ColWidth(0) = 800: CapOb1.CellAlignment = 4: CapOb1.Text = "No."
  CapOb1.Col = 1: CapOb1.ColWidth(1) = 800: CapOb1.CellAlignment = 4: CapOb1.Text = "Obra"
  CapOb1.Col = 2: CapOb1.ColWidth(2) = 800: CapOb1.CellAlignment = 4: CapOb1.Text = "Porcentaje"
  CapOb1.Col = 3: CapOb1.ColWidth(2) = 1200: CapOb1.CellAlignment = 4: CapOb1.Text = "Importe"
  
End Sub

Private Sub Form_Resize()
   Aplper.Width = 4100
   Aplper.Height = (35 * TextHeight("Q"))
   CapOb1.Height = ScaleHeight
   CapOb1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If ultimo.num = 0 Then
    respuesta = MsgBox("Desea archivar los datos", vbYesNo, "Cambio Archivos Obras")
    If respuesta = vbYes Then
       ArcGua_Click
    End If
  End If
End Sub

Private Sub Text1_Change()
    CapOb1.Text = Text1.Text
End Sub

Private Sub Text1_GotFocus()
   SendKeys "{end}"
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
      Case vbKeyEscape
            Text1.Text = valcelant
            
            Rem SendKeys "{end}"
      Case vbKeyBack
            Rem nada
      Case vbKeyClear
        CapOb1.Text = ""
      
      
    End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Select Case CapOb1.Col
        Case 1
            buscaobra
            If selocalizo = 1 Then
               selocalizo = 0
               CapOb1.Text = Text1.Text
               CapOb1.SetFocus
               
               Else
               Rem nada
            End If
        Case 2
            Rem ON ERROR GoTo novale
            If Text1.Text >= 1 And Text1.Text <= 100 Then
                    CapOb1.Text = Text1.Text
                    cantid = (base * Text1.Text / 100)
                    CapOb1.TextMatrix(CapOb1.Row, 3) = Format(cantid, z1)
                    CapOb1.SetFocus
                    vertsum
                    ultimo.num = 0
                    Else
                    MsgBox "valor no permitido"
                    Text1_KeyDown 27, 1
            End If
       End Select
       Exit Sub
   End If
Exit Sub
novale:
         MsgBox "valor no permitido"
         Text1_KeyDown 27, 1

End Sub
Sub buscaobra()
  
   Dim aqui As Integer, respuesta
   If Text1.Text = 9000 Then selocalizo = 1: Exit Sub
   aqui = 0: selocalizo = 0
   For f = 11 To 410: Get 9, f, CATAUX
      If (Val(CATAUX.C1) = 0) And (aqui = 0) Then aqui = f
      If Val(CATAUX.C1) = Text1.Text Then
          MsgBox CATAUX.C1 + " " + CATAUX.C2, vbDefaultButton1
          selocalizo = 1
          Exit For
      End If
   Next f
   If selocalizo = 0 Then
      Entrada.texto.Text = Text1.Text
      Entrada.Caption = "Aplicación nómina"
      Entrada.etiqueta = "No existe nombre de obra" & Chr(13) & Chr(10) & "Use maximo 28 caracteres "
      Entrada.texto.MaxLength = 32
      Entrada.texto.Width = 180 * 32
      Entrada.Show 1
      
      If APLICAR = True Then
        
        Get 9, aqui, CATAUX: CATAUX.C1 = Left(Text1.Text, 4): CATAUX.C2 = ultimo.texto
        
        Put 9, aqui, CATAUX
        Get 9, (aqui + 400), CATAUX: CATAUX.C1 = Text1.Text: CATAUX.C2 = ultimo.texto
        Put 9, (aqui + 400), CATAUX
        Get 9, (aqui + 800), CATAUX: CATAUX.C1 = Text1.Text: CATAUX.C2 = ultimo.texto
        Put 9, (aqui + 800), CATAUX
        APLICAR = False
        selocalizo = 1
        Else
         selocalizo = 0
         Text1_KeyDown 27, 1
      End If
   End If
End Sub



