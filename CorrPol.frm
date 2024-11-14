VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form CorrPol 
   Caption         =   "Visor de Operaciones"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   11496
   Icon            =   "CorrPol.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   11496
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "INICIAR"
      Height          =   372
      Left            =   480
      TabIndex        =   3
      Top             =   6120
      Width           =   1572
   End
   Begin VB.PictureBox Barra 
      Height          =   252
      Left            =   2640
      ScaleHeight     =   204
      ScaleWidth      =   6204
      TabIndex        =   1
      Top             =   6120
      Width           =   6252
      Begin MSComctlLib.ProgressBar Barra1 
         Height          =   252
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   6252
         _ExtentX        =   11028
         _ExtentY        =   445
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Visor 
      Height          =   5172
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   10692
      _ExtentX        =   18860
      _ExtentY        =   9123
      _Version        =   393216
   End
End
Attribute VB_Name = "CorrPol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mens_1 As Integer, Mens_2 As Integer, Mens_3 As Integer
Private Sub Command1_Click()
  Form_Initialize
  Form_Load
End Sub

Private Sub Command1_GotFocus()
  Form_Load
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Command1_Click
End Sub

Private Sub Form_Load()
   Colum
   Locpol
End Sub
Sub Sum_pol(IncA, FinZ)
    Dim R_ini As Long, S_cta As Currency, R_fin As Long, Real As Long
    Dim CTA As Long
    Mens_1 = 0: Mens_2 = 0: Mens_3 = 0
    S_cta = 0: R_ini = 0: R_fin = 0
     Close 3
        Close 2: Open "Catmay" For Random As 2 Len = Len(CATMAY)
        Close 4: Open "Cataux" For Random As 4 Len = Len(CATAUX)
    For r1 = (IncA + 1) To FinZ
            Barra1.Refresh
            Barra1.Value = r
        AIDT = VeroperI.VER1.TextMatrix(r1, 5)
              Select Case AIDT
                     Case "B"
                        If S_cta <> 0 Then
                            MsgBox "Existe un error, no coincide el saldo de los auxiliares con el del mayor"
                            Mens_2 = 1
                        End If
                        CTA = VeroperI.VER1.TextMatrix(r1, 6)
                        Get 2, CTA, CATMAY
                        If IsNumeric(CATMAY.B4) And IsNumeric(CATMAY.B5) Then
                                   R_ini = CATMAY.B4: R_fin = CATMAY.B5
                                   Else
                                   R_ini = 0: R_fin = 0
                        End If
                        impte = VeroperI.VER1.TextMatrix(r1, 4)
                        Rem descr = VeroperI.VER1.TextMatrix(r1, 2)
                        
                        If Val(impte) > 0 Then
                            Rem Visor.AddItem Reg & Chr(9) & Format(CATMAY.B1, "###0") & Chr(9) & "" & Chr(9) & _
                                     CATMAY.B2 & Chr(9) & "" & Chr(9) & _
                                     Format(impte, "###,###,##0.00") _
                                     & Chr(9) & "" & Chr(9) & AIDT & Chr(9) & "" & Chr(9) & descr
                            Mvtodebe = Mvtodebe + impte
                            If R_ini > 0 Then
                                S_cta = impte
                            End If
                            Else
                            Rem Visor.AddItem Reg & Chr(9) & Format(CATMAY.B1, "###0") & Chr(9) & "" & Chr(9) & _
                                     CATMAY.B2 & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format(impte, "###,###,##0.00") _
                                     & Chr(9) & AIDT & Chr(9) & descr
                            Mvtohaber = Mvtohaber + impte
                            If R_ini > 0 Then
                                S_cta = impte
                            End If

                        End If
                                     
                
                        Case "C"
                          Rem NADA
                        Real = VeroperI.VER1.TextMatrix(r1, 1)
                        SUBCTA = VeroperI.VER1.TextMatrix(r1, 6)
                        Get 2, Val(Real), CATMAY
                        impte = VeroperI.VER1.TextMatrix(r1, 4)
                        Rem descr = VeroperI.VER1.TextMatrix(r1, 2)
                        If Real >= R_ini And Real <= R_fin Then
                                Get 4, Real, CATAUX
                                Else
                                MsgBox "El auxiliar esta fuera de rango de la cuenta "
                                Mens_3 = 1
                        End If
                            S_cta = S_cta - impte
                            Rem Visor.AddItem Reg & Chr(9) & "" & Chr(9) & Format(CATAUX.C1, "###0") & Chr(9) & _
                                     (" " + CATAUX.C2) & Chr(9) & Format(impte, "###,###,##0.00") _
                                     & Chr(9) & "" & Chr(9) & "" & Chr(9) & AIDT & Chr(9) & _
                                     descr
                      Case "A"
                      Rem NADA
           End Select
    Next r1
    SALDO = Mvtodebe + Mvtohaber
    If SALDO <> 0 Then
       MsgBox " las sumas no son iguales"
       Mens_1 = 1
    End If
    If (Mens_1 = 1) Or (Mens_2 = 1) Or (Mens_3 = 1) Then
            verpoliza R_ini, R_fin
    End If
End Sub
Sub verpoliza(IncA As Long, FinZ As Long)
    
     Dim Real As Long, AIDT As String, CTA As Long
     Dim SUBCTA As Long
        Visor.Rows = 1
        Visor.Row = 0
        Close 3
        Close 2: Open "Catmay" For Random As 2 Len = Len(CATMAY)
        Close 4: Open "Cataux" For Random As 4 Len = Len(CATAUX)
        Mvtodebe = 0: Mvtohaber = 0
        For r1 = (IncA + 1) To FinZ
              Reg = VeroperI.VER1.TextMatrix(r1, 0)
              AIDT = VeroperI.VER1.TextMatrix(r1, 5)
              Select Case AIDT
                     Case "B"
                        CTA = VeroperI.VER1.TextMatrix(r1, 6)
                        Get 2, CTA, CATMAY
                        
                        impte = VeroperI.VER1.TextMatrix(r1, 4)
                        descr = VeroperI.VER1.TextMatrix(r1, 2)
                        If Val(impte) > 0 Then
                            Visor.AddItem Reg & Chr(9) & Format(CATMAY.B1, "###0") & Chr(9) & "" & Chr(9) & _
                                     CATMAY.B2 & Chr(9) & "" & Chr(9) & _
                                     Format(impte, "###,###,##0.00") _
                                     & Chr(9) & "" & Chr(9) & AIDT & Chr(9) & "" & Chr(9) & descr
                            Mvtodebe = Mvtodebe + impte
                            Else
                            Visor.AddItem Reg & Chr(9) & Format(CATMAY.B1, "###0") & Chr(9) & "" & Chr(9) & _
                                     CATMAY.B2 & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format(impte, "###,###,##0.00") _
                                     & Chr(9) & AIDT & Chr(9) & descr
                            Mvtohaber = Mvtohaber + impte
                        End If
                                     
                
                        Case "C"
                        
                        Real = VeroperI.VER1.TextMatrix(r1, 1)
                        SUBCTA = VeroperI.VER1.TextMatrix(r1, 6)
                        Rem Get 2, Val(Real), CATMAY
                        impte = VeroperI.VER1.TextMatrix(r1, 4)
                        descr = VeroperI.VER1.TextMatrix(r1, 2)
     
                        Get 4, Real, CATAUX
                          
                            Visor.AddItem Reg & Chr(9) & "" & Chr(9) & Format(CATAUX.C1, "###0") & Chr(9) & _
                                     (" " + CATAUX.C2) & Chr(9) & Format(impte, "###,###,##0.00") _
                                     & Chr(9) & "" & Chr(9) & "" & Chr(9) & AIDT & Chr(9) & _
                                     descr
                
                        Case "A"
                        Rem NADA
                  End Select
        Next r1
         Rem Close 2, 4
        Visor.AddItem "" & Chr(9) & "" & Chr(9) & _
                                     "Sumas Iguales" & Chr(9) & "" & Chr(9) & Format(Mvtodebe, "###,###,##0.00") & Chr(9) & Format(Mvtohaber, "###,###,##0.00") _
                                     & Chr(9) & "" & Chr(9) & "" & Chr(9) & ""
                
        
        Rem Visor.Height = (Visor.Rows + 4) * (Visor.CellHeight)
        Rem If Visor.Height > TopeBorrego Then Visor.Height = TopeBorrego
            Rem Form_Resize
       Rem  End If
       Rem InputBox "PRESIONA CUALQUIER TECLA PARA CONTINUAR"
       Rem Close 3
End Sub

Sub Colum()
    Visor.Clear
    Visor.Cols = 10
    Visor.FixedCols = 1
    Visor.Rows = 1
    Visor.Row = 0:
    Visor.Col = 0: Visor.ColWidth(0) = 600: Visor.CellFontBold = True: Visor.CellAlignment = 4: Visor.Text = "Rgto."
    Visor.Col = 1: Visor.ColWidth(1) = 600: Visor.CellFontBold = True: Visor.CellAlignment = 4: Visor.Text = "Cta."
    Visor.Row = 0: Visor.Col = 2: Visor.ColWidth(2) = 600: Visor.CellFontBold = True: Visor.CellAlignment = 4: Visor.Text = "SubCta."
    Visor.Row = 0: Visor.Col = 3: Visor.ColWidth(3) = 2700: Visor.CellFontBold = True: Visor.CellAlignment = 4: Visor.Text = "Nombre."
    Visor.Row = 0: Visor.Col = 4: Visor.ColWidth(4) = 1200: Visor.CellFontBold = True: Visor.CellAlignment = 4: Visor.Text = "Parcial."
    Visor.Row = 0: Visor.Col = 5: Visor.ColWidth(5) = 1200: Visor.CellFontBold = True: Visor.CellAlignment = 4: Visor.Text = "Debe."
    Visor.Row = 0: Visor.Col = 6: Visor.ColWidth(6) = 1200: Visor.CellFontBold = True: Visor.CellAlignment = 4: Visor.Text = "Haber."
    Visor.Row = 0: Visor.Col = 7: Visor.ColWidth(7) = 600: Visor.CellFontBold = True: Visor.CellAlignment = 4: Visor.Text = "Cl"
    Visor.Row = 0: Visor.Col = 8: Visor.ColWidth(8) = 2700: Visor.CellFontBold = True: Visor.CellAlignment = 4: Visor.Text = "Concepto."
    Visor.Row = 0: Visor.Col = 9: Visor.ColWidth(9) = 600: Visor.CellFontBold = True: Visor.CellAlignment = 4: Visor.Text = "others1"
End Sub
Sub Locpol()
 Dim InA As Long, Fza As Long
 InA = 0
           Barra1.Scrolling = 0
           Barra1.Max = VeroperI.VER1.Rows - 1
           Barra1.Min = 1
           Barra1.Value = 1
           Barra1.Refresh
           Rem  Barra1.Value = RGTRO
 For r = 1 To VeroperI.VER1.Rows - 1
           Barra1.Refresh
           Barra1.Value = r
        If (VeroperI.VER1.TextMatrix(r, 5) = "A") Then
           If InA = 0 Then
                InA = r
                Fza = 0
              Else
               Fza = r
               Sum_pol InA, Fza
               Rem verpoliza InA, Fza
           End If
             'numero = VeroperI.VER1.TextMatrix(r, 1)
             'num1 = Format(numero, "#####0")
             'descr = VeroperI.VER1.TextMatrix(r, 2)
             Rem num1 = String(6 - Len(num1), " ") + num1
             Rem muestra = r & Chr(9) & Chr(9) & Chr(9) & " " + descr & Chr(9) & Str(r)
             Rem Visor.AddItem muestra
            
        End If
        
 Next r
End Sub

