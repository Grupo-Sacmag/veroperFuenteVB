VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Poliza 
   Caption         =   "Registro de operaciones"
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   8070
   Icon            =   "CG.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Apl1 
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6588
      _Version        =   393216
      Rows            =   3
      Cols            =   3
      FixedRows       =   2
      FixedCols       =   2
      BackColorBkg    =   -2147483637
      Appearance      =   0
   End
   Begin VB.Menu ApArc 
      Caption         =   "&Archivo"
      Begin VB.Menu ApRec 
         Caption         =   "&Recalcular"
      End
      Begin VB.Menu Aplsep1 
         Caption         =   "-"
      End
      Begin VB.Menu AplCorr 
         Caption         =   "&Corregir aplicación"
         Shortcut        =   ^A
      End
      Begin VB.Menu AplSep2 
         Caption         =   "-"
      End
      Begin VB.Menu AplImp 
         Caption         =   "Imprimir"
         Begin VB.Menu AplImpRe 
            Caption         =   "&Resumen"
         End
         Begin VB.Menu AplImpTot 
            Caption         =   "&Total"
         End
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "&Edicion"
      Begin VB.Menu EditCop 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu EditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu EditSelT 
         Caption         =   "&Seleccionar Todo"
      End
   End
   Begin VB.Menu Orden 
      Caption         =   "&Ordenar"
      Begin VB.Menu OrdAlf 
         Caption         =   "&Alfabetico"
      End
      Begin VB.Menu ORSEP1 
         Caption         =   "-"
      End
      Begin VB.Menu OrdNum 
         Caption         =   "&Numerico"
      End
   End
   Begin VB.Menu AplPol 
      Caption         =   "&Generar Poliza"
      Begin VB.Menu AplVerP 
         Caption         =   "Aplicacion Poliza"
      End
   End
End
Attribute VB_Name = "Poliza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Directorio As String, personas As String, lnomina As String, DirCost As String
Dim cm As Long, Dm  As Long, cont_obr As Integer, col_obra(100) As Integer, gm As Long
Dim i As Integer, f As Integer, cantid As Currency, cont_reng As Integer, cantidad
Sub sumahor()
  For f = 2 To Apl1.Rows - 1
   cantid = 0
   For r = 3 To Apl1.Cols - 1
      If Apl1.TextMatrix(f, r) <> "" Then
            cantid = cantid + Apl1.TextMatrix(f, r)
       End If
   Next r
    Apl1.TextMatrix(f, 2) = Format(cantid, z1)
    
  Next f
End Sub
Sub sumavert()
  For f = 3 To Apl1.Cols - 1
   cantid = 0
   For r = 2 To Apl1.Rows - 2
      If Apl1.TextMatrix(r, f) <> "" Then
            cantid = cantid + Apl1.TextMatrix(r, f)
       End If
   Next r
    Apl1.TextMatrix((Apl1.Rows - 1), f) = Format(cantid, z1)
    
  Next f
End Sub

Private Sub Apl1_DblClick()
   Get 8, Apl1.TextMatrix(Apl1.Row, 0), maestro
   base = Apl1.TextMatrix(Apl1.Row, 2)
   tar = Apl1.TextMatrix(Apl1.Row, 0)
   Aplper.Show 1
   Apl1.Col = 0: Apl1.CellForeColor = vbBlue
   Apl1.Col = 1: Apl1.CellForeColor = vbBlue
   Apl1.Col = 2: Apl1.CellForeColor = vbBlue
End Sub

Private Sub Apl1_EnterCell()
  
 If (Apl1.Row > 1) And (Apl1.Col > 1) Then
  Apl1.CellBackColor = vbYellow
 End If

End Sub



Private Sub Apl1_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then Apl1_DblClick
End Sub

Private Sub Apl1_LeaveCell()
   If Apl1.Rows > 1 Then
    If Apl1.Col > 1 And Apl1.Row > 1 Then
         Apl1.CellBackColor = vbWhite
    End If
 End If
End Sub
    

Private Sub AplCorr_Click()
  Apl1_DblClick
End Sub
Sub tituloApl()
   Dim letrero As String
   centrar ancho2, empresa.name, Printer.Width
   Printer.CurrentX = ancho2
   Printer.Print empresa.name
   letrero = "Aplicación " + Form8.Label7.Caption
   centrar ancho2, letrero, Printer.Width
   Printer.CurrentX = ancho2
   Printer.Print letrero
   Printer.Print
   Printer.Line (2800, Printer.CurrentY)-(9200, Printer.CurrentY + 50), , BF
   Printer.CurrentX = 3500: Printer.Print "O  b  r  a ";
   Printer.CurrentX = 8300: Printer.Print "Importe"
   Printer.Line (2800, Printer.CurrentY)-(9200, Printer.CurrentY + 50), , BF
   Printer.Print

End Sub
Sub PieApl()
   Printer.Line (2800, Printer.CurrentY)-(9200, Printer.CurrentY + 10), , BF
   Printer.CurrentX = 3500: Printer.Print "S  u  m  a ";
   colocar ancho2, Apl1.TextMatrix(Apl1.Rows - 1, 2), z1
   Printer.CurrentX = 8200 + ancho2
   Printer.Print Format(Apl1.TextMatrix(Apl1.Rows - 1, 2), z1)
   Printer.Line (8200, Printer.CurrentY)-(9200, Printer.CurrentY), , B
   Printer.Line (8200, Printer.CurrentY + 20)-(9200, Printer.CurrentY + 20), , B

End Sub
Private Sub AplImpRe_Click()
     tituloApl
     For i = 3 To Apl1.Cols - 1
     If Apl1.TextMatrix(Apl1.Rows - 1, i) <> "" Then
        colocar ancho2, Apl1.TextMatrix(0, i), "###0"
        Printer.CurrentX = 2800 + ancho2
        Printer.Print Format(Apl1.TextMatrix(0, i), "####0");
        Printer.CurrentX = 3200
        Printer.Print Apl1.TextMatrix(1, i);
        colocar ancho2, Apl1.TextMatrix(Apl1.Rows - 1, i), z1
        Printer.CurrentX = 8200 + ancho2
        Printer.Print Format(Apl1.TextMatrix(Apl1.Rows - 1, i), z1)
     End If
   Next i
   PieApl
   Printer.EndDoc
End Sub
Sub TitulObr()
Dim letrero As String
centrar ancho2, empresa.name, Printer.Width
Printer.CurrentX = ancho2
If Printer.CurrentY < 800 Then Printer.Print empresa.name
letrero = RTrim(Apl1.TextMatrix(0, f)) + " " + RTrim(Apl1.TextMatrix(1, f)) + "  " + RTrim(Form8.Label7.Caption)
centrar ancho2, letrero, Printer.Width
Printer.CurrentX = ancho2
Printer.Print letrero
Printer.Line (2800, Printer.CurrentY)-(9200, Printer.CurrentY + 20), , BF
Printer.CurrentX = 3500: Printer.Print "E m p l e a d o ";
Printer.CurrentX = 8300: Printer.Print "Importe"
Printer.Line (2800, Printer.CurrentY)-(9200, Printer.CurrentY + 20), , BF
Printer.Print
End Sub
Sub PieAplTot()
   Printer.Line (2800, Printer.CurrentY)-(9200, Printer.CurrentY + 10), , BF
   Printer.CurrentX = 3500: Printer.Print "S  u  m  a ";
   valor$ = Format(cantid, z1)
   colocar ancho2, valor$, z1
   Printer.CurrentX = 8200 + ancho2
   Printer.Print Format(cantid, z1)
   Printer.Line (8200, Printer.CurrentY)-(9200, Printer.CurrentY), , B
   Printer.Line (8200, Printer.CurrentY + 20)-(9200, Printer.CurrentY + 20), , B
   Printer.Print
           
End Sub

Private Sub AplImpTot_Click()
    For f = 3 To Apl1.Cols - 1
       TitulObr
       cantid = 0
       For i = 2 To Apl1.Rows - 2
           If Apl1.TextMatrix(i, f) <> "" Then
             colocar ancho2, Apl1.TextMatrix(i, 0), "###0"
             Printer.CurrentX = 2800 + ancho2
             Printer.Print Format(Apl1.TextMatrix(i, 0), "####0");
             Printer.CurrentX = 3200
             Printer.Print Apl1.TextMatrix(i, 1);
             colocar ancho2, Apl1.TextMatrix(i, f), z1
             Printer.CurrentX = 8200 + ancho2
             Printer.Print Format(Apl1.TextMatrix(i, f), z1)
             cantid = cantid + Apl1.TextMatrix(i, f)
             If Printer.CurrentY > (Printer.Height - 1000) Then
                 Printer.NewPage
                 TitulObr
             End If
           End If
       Next i
       PieAplTot
   Next f
   Printer.EndDoc
End Sub

Private Sub AplVerP_Click()
  Rem Poliza1.Show
  AplNom.Show
End Sub

Private Sub ApRec_Click()
   Apl1.Clear
   Form_Load
End Sub

Private Sub EditCop_Click()
    Dim Temporal1
    Clipboard.Clear
   Rem Clipboard.SetText Clipboard.GetText + Poliza1.Caption & Chr(13)
   Rem Clipboard.SetText Clipboard.GetText + Label1.Caption & Chr(13)
   Rem Apl1.RowSel = Apl1.Rows - 1
   Rem Apl1.ColSel = Apl1.Cols - 1
   Temporal1 = Temporal1 + Apl1.TextMatrix(0, 0) & Chr(9) & Apl1.TextMatrix(0, 1) & Chr(9)
   For f = Apl1.Col To Apl1.ColSel
         Temporal1 = Temporal1 + Apl1.TextMatrix(0, f) + " " + Apl1.TextMatrix(1, f) & Chr(9)
   Next f
        Clipboard.SetText Temporal1 & Chr(13)
   For i = Apl1.Row To Apl1.RowSel
           Temporal1 = Temporal1 + Apl1.TextMatrix(i, 0) & Chr(9) + Apl1.TextMatrix(i, 1) & Chr(9)
      For f = Apl1.Col To Apl1.ColSel
            Temporal1 = Temporal1 + Apl1.TextMatrix(i, f) & Chr(9)
      Next f
      Clipboard.SetText Temporal1 & Chr(13)
   Next i
   difer = Apl1.RowSel - Apl1.Row
    
End Sub

Private Sub EditSelT_Click()
    Clipboard.Clear
    Apl1.RowSel = Apl1.Rows - 1
    Apl1.ColSel = Apl1.Cols - 1

End Sub

Private Sub Form_Load()
    cont_reng = 2
    abre_arch
    gene_col
    desc_nomina
    sumavert
    sumahor
    negritas
    Apl1.Row = 2: Apl1.Col = 2
    Apl1_LeaveCell
    Apl1_EnterCell
End Sub
Sub negritas()
  For i = 0 To Apl1.Cols - 1
      Apl1.Row = 0
      Apl1.Col = i
      ver = Apl1.Text
      Apl1.CellFontBold = True
      Apl1.CellFontSize = 2
      Apl1.Text = Apl1.Text
      Apl1_LeaveCell
      Apl1_EnterCell
      Apl1.Row = 1
      Apl1.Col = i
      ver = Apl1.Text
      Apl1.CellFontSize = 2
      Apl1.CellFontBold = True
      Apl1.Text = ver
      Apl1_LeaveCell
      Apl1_EnterCell
   
  Next i

End Sub
Sub desc_nomina()
  Apl1.Row = 1
  For r = 1 To Form8.ConNom1.Rows - 3:
  If Form8.ConNom1.TextMatrix(r, 11) > 0 Then
        Apl1.Row = Apl1.Row + 1
        regtro = Form8.ConNom1.TextMatrix(r, 0)
        Get 2, regtro, personal
        Apl1.TextMatrix(Apl1.Row, 0) = Format(regtro, "####0")
        Nombre = " " + LTrim(RTrim(personal.ape1)) + " " + RTrim(personal.nom)
        Apl1.TextMatrix(Apl1.Row, 1) = Nombre
        Get 8, regtro, maestro
        convierte
        base = Form8.ConNom1.TextMatrix(r, 11)
        For i = 1 To 20
           If obra(1) = 0 Then
             obra(1) = 9000:
             porcentaje(1) = 100
             grabamaestro
             Put 8, regtro, maestro
           End If
           If obra(i) > 0 Then
                siexiste = False
                For f = 3 To Apl1.Cols - 1
                    If Apl1.TextMatrix(0, f) = obra(i) Then
                        cantid = base * porcentaje(i) / 100
                        Apl1.TextMatrix(Apl1.Row, f) = Format(cantid, z1)
                        Exit For
                    End If
                Next f
           End If
              
        Next i
     End If
  Next r
   
End Sub
Sub gene_col()
  
  For r = 1 To Form8.ConNom1.Rows - 3
  
     If Form8.ConNom1.TextMatrix(r, 11) > 0 Then
        cont_reng = cont_reng + 1
        regtro = Form8.ConNom1.TextMatrix(r, 0)
        Get 8, regtro, maestro
        agrega_obra
     End If
  Next r
  Apl1.Cols = cont_obr + 3
  Apl1.Rows = cont_reng + 1
  ordenar
  Apl1.Row = 0
  Apl1.ColWidth(0) = 800
  Apl1.ColWidth(1) = 2800
  
  Apl1.Col = 0: Apl1.ColWidth(0) = 800: Apl1.CellAlignment = 4: Apl1.Text = "No."
  Apl1.Col = 1: Apl1.ColWidth(1) = 2800: Apl1.CellAlignment = 4: Apl1.Text = "Nombre"
  Apl1.Col = 2: Apl1.ColWidth(2) = 1200: Apl1.CellAlignment = 4: Apl1.Text = "Sumas"

  For i = 1 To cont_obr
     Apl1.Col = i + 2: Apl1.ColWidth(i + 2) = 1200: Apl1.CellAlignment = 4
     Apl1.Text = Format(col_obra(i), "####0")
     BuscNom
  Next i
  
End Sub
Sub abre_arch()
  Rem directorio = "\SUP2002\NOMINA\"
  Rem personas = directorio + "personal.dno"
  Rem dirobr = directorio + "maestro.dno"
  Rem lnomina = directorio + "MAY12002.NOM"
  DirCost = dir_obras + "Cataux"
  
  Rem Open personas For Random As 1 Len = Len(personal)
  Rem cm = LOF(1) / Len(personal)
  Rem Open lnomina For Random As 2 Len = Len(nomina)
  Rem dm = LOF(2) / Len(nomina)
  Rem Open dirobr For Random As 3 Len = Len(maestro)
  Rem em = LOF(3) / Len(maestro)
  Close 9: Open DirCost For Random As 9 Len = Len(CATAUX)
  gm = LOF(9) / Len(CATAUX)
End Sub

Sub BuscNom()
   Rem Apl1.Row = 1
   For f = 11 To 410: Get 9, f, CATAUX
      If Val(CATAUX.C1) = col_obra(i) Then
          
          Apl1.CellFontSize = 4
          Apl1.ColAlignment(i + 2) = 3
          Apl1.TextMatrix(1, i + 2) = RTrim(Mid(CATAUX.C2, 1, 28))
               
          
          Exit For
      End If
   Next f
   If col_obra(i) = 9000 Then
          
          Apl1.CellFontWidth = 4
          Apl1.ColAlignment(i + 2) = 3
          Apl1.TextMatrix(1, (i + 2)) = "Grales"
          
   End If
   Rem Apl1.Row = 2
   Rem Apl1.Col = 2
End Sub
Sub agrega_obra()
  
  convierte
  For i = 1 To 20
    If obra(1) = 0 Then obra(1) = 9000: porcentaje(1) = 100
    If obra(i) > 9999 Then obra(i) = 0
    If obra(i) > 0 Then
       siexiste = False
       For f = 1 To cont_obr
          If col_obra(f) = obra(i) Then
             siexiste = True
             Exit For
          End If
       Next f
       If siexiste = False Then
           cont_obr = cont_obr + 1
           col_obra(cont_obr) = obra(i)
       End If
    
    End If
   Next i
End Sub
Sub ordenar()
  A = 0
  Do Until A > cont_obr
    A = A + 1: aa = A
    qq = col_obra(A): A2 = A
    For ji = A2 To cont_obr
        If col_obra(ji) < qq Then
            qq = col_obra(ji)
            aa = ji
         End If
    
  Next ji
  col_obra(aa) = col_obra(A): col_obra(A) = qq
  Loop
        
End Sub

Private Sub Form_Resize()
    Apl1.Height = ScaleHeight * 0.9
    Apl1.Width = ScaleWidth
End Sub

Private Sub OrdAlf_Click()
    Apl1_LeaveCell
    Apl1.Row = 2
    Apl1.Col = 1
    Apl1.RowSel = Apl1.Rows - 2
    Apl1.Sort = 1
    Apl1.Col = 2
    Apl1.Row = 2

    Apl1_LeaveCell
    Apl1.SetFocus
  
End Sub

Private Sub OrdNum_Click()
    Apl1_LeaveCell
    Apl1.Row = 2
    Apl1.Col = 0
    Apl1.RowSel = Apl1.Rows - 2
    Apl1.Sort = 1
    Apl1.Col = 2
    Apl1.Row = 2
    Apl1_LeaveCell
    Apl1.SetFocus

End Sub

