VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form VDatos 
   Caption         =   "CorrDatos"
   ClientHeight    =   4830
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MDatos 
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   6165
      _Version        =   393216
   End
   Begin VB.Menu Arc 
      Caption         =   "&Archivo"
      Begin VB.Menu ArcAbr 
         Caption         =   "&Abrir"
      End
      Begin VB.Menu ArchGuarda 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu ArcSep1 
         Caption         =   "-"
      End
      Begin VB.Menu ArcSAle 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "VDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim F_Aum As Double

Private Sub ArcAbr_Click()
   
  On Err GoTo ErrHandler
            CommonDialog1.Flags = cdlOFNHideReadOnly
            
                    CommonDialog1.FileName = "DAT*.*"
                    CommonDialog1.Filter = "Archivos de (DAT*.*)|DAT*.*"
                               
            CommonDialog1.ShowOpen
            If CommonDialog1.FileName <> "" Then
               CARGADATOS
                    Else
                    nombrearchivo = ""
            End If
ErrHandler:
   Close
   Exit Sub

End Sub
Sub Guarda()
    Open CommonDialog1.FileName For Random As 1 Len = Len(Datos)
   CM = LOF(1) / Len(Datos)
   For r = 1 To CM: Get 1, r, Datos
    MDatos.Row = r
    MDatos.Col = 1: Datos.D1 = MDatos.Text
    MDatos.Col = 2: Datos.D2 = MDatos.Text
    MDatos.Col = 3: Datos.D3 = MDatos.Text
    MDatos.Col = 4: Datos.No_arch = MDatos.Text
    MDatos.Col = 5: Datos.a_o = Str(MDatos.Text)
    MDatos.Col = 6: Datos.others1 = MDatos.Text
    MDatos.Col = 7: Datos.UltimaPol = (MDatos.Text)
    MDatos.Col = 8: Datos.UltimoReg = (MDatos.Text)
    MDatos.Col = 9: Datos.others = MDatos.Text
    Put 1, r, Datos
 Next r
 Close
End Sub
Sub Colum_nas()
    MDatos.Clear
    MDatos.Cols = 10
    MDatos.FixedCols = 1
    MDatos.Rows = 1
    MDatos.Row = 0:
    MDatos.Col = 0: MDatos.ColWidth(0) = 1200: MDatos.CellFontBold = True: MDatos.CellAlignment = 4: MDatos.Text = "Mes"
    MDatos.Col = 1: MDatos.ColWidth(1) = 3600: MDatos.CellFontBold = True: MDatos.CellAlignment = 4: MDatos.Text = "EMPRESA"
    MDatos.Row = 0: MDatos.Col = 2: MDatos.ColWidth(2) = 600: MDatos.CellFontBold = True: MDatos.CellAlignment = 4: MDatos.Text = "D2"
    MDatos.Row = 0: MDatos.Col = 3: MDatos.ColWidth(3) = 600: MDatos.CellFontBold = True: MDatos.CellAlignment = 4: MDatos.Text = "D3"
    MDatos.Row = 0: MDatos.Col = 4: MDatos.ColWidth(4) = 1200: MDatos.CellFontBold = True: MDatos.CellAlignment = 4: MDatos.Text = "ARCHIVO"
    MDatos.Row = 0: MDatos.Col = 5: MDatos.ColWidth(5) = 800: MDatos.CellFontBold = True: MDatos.CellAlignment = 4: MDatos.Text = "Año"
    MDatos.Row = 0: MDatos.Col = 6: MDatos.ColWidth(6) = 800: MDatos.CellFontBold = True: MDatos.CellAlignment = 4: MDatos.Text = "others"
    MDatos.Row = 0: MDatos.Col = 7: MDatos.ColWidth(7) = 1200: MDatos.CellFontBold = True: MDatos.CellAlignment = 4: MDatos.Text = "POLIZA"
    MDatos.Row = 0: MDatos.Col = 8: MDatos.ColWidth(8) = 1200: MDatos.CellFontBold = True: MDatos.CellAlignment = 4: MDatos.Text = "REGISTRO"
    MDatos.Row = 0: MDatos.Col = 9: MDatos.ColWidth(9) = 800: MDatos.CellFontBold = True: MDatos.CellAlignment = 4: MDatos.Text = "others1"
End Sub
Sub CARGADATOS()
   Dim AFecha As Double, Trasp
   Close
   Open CommonDialog1.FileName For Random As 1 Len = Len(Datos)
   CM = LOF(1) / Len(Datos)
   If CM = 0 Then CM = 1
   For r = 1 To CM: Get 1, r, Datos
    AFecha = Val(Datos.a_o)
    Trasp = ""
    
    'Trasp = Datos.D1 & Chr(9) & Datos.UltimaPol _
                     & Chr(9) & Datos.UltimoReg & Chr(9) & Datos.D2 _
                     & Chr(9) & Datos.a_o & Chr(9) & Datos.D3 _
                     & Chr(9) & Datos.No_arch & Chr(9) & Datos.others
    MDatos.AddItem Trasp
    MDatos.Row = r
    MDatos.Col = 0: MDatos.Text = Mm(r)
    MDatos.Col = 1: MDatos.Text = Datos.D1
    MDatos.Col = 2: MDatos.Text = Datos.D2
    MDatos.Col = 3: MDatos.Text = Datos.D3
    MDatos.Col = 4: MDatos.Text = Datos.No_arch
    MDatos.Col = 5: MDatos.Text = Val(Datos.a_o)
    MDatos.Col = 6: MDatos.Text = Datos.others1
    'MDatos.Col = 7: MDatos.Text = Val(Datos.UltimaPol)
    MDatos.Col = 7: MDatos.Text = (Datos.UltimaPol)
    MDatos.Col = 8: MDatos.Text = (Datos.UltimoReg)
    MDatos.Col = 9: MDatos.Text = Datos.others
    'D1 As String * 64
    'D2 As String * 60
    'D3 As String * 45
    'No_arch As String * 15
    'a_o As String * 5
    'others1  As String * 25
    'UltimaPol As String * 5
    'UltimoReg As String * 5
    'others As String * 12
    
 Next r
End Sub

Private Sub ArchGuarda_Click()
   
  Respuesta = MsgBox("Se modificara el archivo de datos", vbYesNo)
  If Respuesta = vbYes Then
      Guarda
      MsgBox "Archivo Guardado"
      Else
      MsgBox "El archivo permanece igual"
  End If
End Sub

Private Sub ArcSAle_Click()
  Unload VDatos
End Sub

Private Sub Form_Load()
  Colum_nas
End Sub

Private Sub Form_Resize()
   If VDatos.WindowState <> 1 Then
      MDatos.Height = ScaleHeight - 200
      MDatos.Width = ScaleWidth - 600
      F_Aum = (MDatos.Width - 600) / 9200
      ColDfn
   End If
End Sub
Sub ColDfn()
    MDatos.FontWidth = 3 * F_Aum
    MDatos.ColWidth(0) = 800 * F_Aum
    MDatos.ColWidth(1) = 2300 * F_Aum
    MDatos.ColWidth(2) = 1500 * F_Aum
    MDatos.ColWidth(3) = 1500 * F_Aum
    MDatos.ColWidth(4) = 1500 * F_Aum
    MDatos.ColWidth(5) = 1500 * F_Aum
    MDatos.ColWidth(0) = 1200 * F_Aum
    MDatos.ColWidth(1) = 3600 * F_Aum
    MDatos.ColWidth(2) = 600 * F_Aum
    MDatos.ColWidth(3) = 600 * F_Aum
    MDatos.ColWidth(4) = 1200 * F_Aum
    MDatos.ColWidth(5) = 800 * F_Aum
    MDatos.ColWidth(6) = 800 * F_Aum
    MDatos.ColWidth(7) = 1200 * F_Aum
    MDatos.ColWidth(8) = 1200 * F_Aum
    MDatos.ColWidth(9) = 1200 * F_Aum
    
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
     
    Select Case KeyAscii
       Case 27
        Text1.Text = valcelant
        MDatos.Text = valcelant
        MDatos.SetFocus

       Case 13
        Select Case (MDatos.Col)
          Case 1, 4, 5, 7, 8
            MDatos.Text = Text1.Text
            MDatos.SetFocus
          'Case 4
                     
            'MDatos.Text = Format(Text1.Text, z1)
            'MDatos.SetFocus
         'Case 5
           'If Option1 = True Then
            'MsgBox "Esta columna no se puede modificar "
            'Text1.Text = valcelant
            'MDatos.Text = valcelant
            'Else
            'MDatos.Text = Text1.Text
            'MDatos.SetFocus
            'End If
        End Select
       Case Else
        MDatos.Text = Text1.Text
 
    End Select
  End Sub
Private Sub Text1_Change()
     If Text1.Text = Chr(27) Then Text1.Text = valcelant
     
     MDatos.Text = Text1.Text
     
End Sub
Private Sub Text1_GotFocus()
    SendKeys "{end}"
    
End Sub
Private Sub MDATOS_EnterCell()
   If MDatos.Col > 0 And MDatos.Row > 0 Then
       MDatos.CellBackColor = vbCyan
       Text1.Text = MDatos.Text
       
       
   End If
End Sub
Private Sub MDATOS_LeaveCell()
    If MDatos.Col > 0 And MDatos.Row > 0 Then
        MDatos.CellBackColor = vbWhite
        
    End If

  End Sub

Private Sub MDATOS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MDatos.RowSel = MDatos.Row
End Sub

Private Sub MDATOS_RowColChange()
     
     If MDatos.Text <> "" Then valcelant = MDatos.Text
     
End Sub

Private Sub MDATOS_KeyDown(KeyCode As Integer, Shift As Integer)
      Select Case KeyCode
            Case vbKeyDelete
              'respuesta = MsgBox("DESEA ELIMINAR ESTE RENGLON ", vbYesNo + vbCritical + vbDefaultButton2)
              'If respuesta = vbYes Then
                'MDATOS.RemoveItem MDATOS.Row
                Rem MDATOS.Text = ""
                
              'End If
                   
    
            If (MDatos.RowSel > MDatos.Row) Or (MDatos.ColSel > MDatos.Col) Then
              Text1.Text = ""
              MDatos.Text = ""
              
            End If
                Text1.Text = ""
                MDatos.Text = ""
                
            Case vbKeyF2
               
                If MDatos.Text <> "" Then valcelant = MDatos.Text
                
                Text1.Text = MDatos.Text
                Text1.SetFocus
               
       End Select
End Sub

Private Sub MDATOS_KeyPress(KeyAscii As Integer)
        MDatos.SelectionMode = flexSelectionFree
        
    
        If MDatos.Text <> "" Then valcelant = MDatos.Text
           Text1.Text = Chr(KeyAscii)
           
           Text1.SetFocus
End Sub


