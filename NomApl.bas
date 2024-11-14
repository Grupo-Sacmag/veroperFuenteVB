Attribute VB_Name = "NomApl"
Type per
    nom As String * 20
    ape1 As String * 20
    ape2 As String * 20
    rfc As String * 18
    imss As String * 18
    fal As String * 12
    fab As String * 12
    ingr As Currency
    viat As Currency
    otras As Currency
    integrado As Currency
 End Type
  Type basini
     datoarch As String * 64
 End Type

 Type nom
     dias As Currency
     hsnor As Currency
     hs_no As Currency
     hsdbl As Currency
     hs_db As Currency
     hstri As Currency
     hs_tr As Currency
     ispt As Currency
     crdsal As Currency
     imss As Currency
     sueldo As Currency
     hs_nor As Currency
     hs_dbl As Currency
     hs_tri As Currency
     viaticos As Currency
     pvac As Currency
     otras As Currency
     aguin As Currency
     ptu As Currency
     exentos As Currency
     prestamos As Currency
     fonacot As Currency
     telefono As Currency
     otraded As Currency
  End Type
 Type art
     liminf As Currency
     limsup As Currency
     cuotaf As Currency
     porcsl As Currency
  End Type
  Type subs
     liminfs As Currency
     limsups As Currency
     cuotafs As Currency
     porcsls As Currency
  End Type
   Type cred
     crede As Currency
     crea As Currency
     cresam As Currency
  End Type
  
  Type empre
       name As String * 60
       ao As Integer
       sm As Currency
       psub As Currency
       fecha As String * 14
  End Type
    Type ob
    O_1 As Integer
    por_1 As Integer
    im_1 As Currency
    O_2 As Integer
    por_2 As Integer
    im_2 As Currency
    O_3 As Integer
    por_3 As Integer
    im_3 As Currency
    O_4 As Integer
    por_4 As Integer
    im_4 As Currency
    O_5 As Integer
    por_5 As Integer
    im_5 As Currency
    O_6 As Integer
    por_6 As Integer
    im_6 As Currency
    O_7 As Integer
    por_7 As Integer
    im_7 As Currency
    O_8 As Integer
    por_8 As Integer
    im_8 As Currency
    O_9 As Currency
    por_9 As Integer
    im_9 As Currency
    O_10 As Integer
    por_10 As Integer
    im_10 As Currency
    O_11 As Integer
    por_11 As Integer
    im_11 As Currency
    O_12 As Integer
    por_12 As Integer
    im_12 As Currency
    O_13 As Integer
    por_13 As Integer
    im_13 As Currency
    O_14 As Integer
    por_14 As Integer
    im_14 As Currency
    O_15 As Integer
    por_15 As Integer
    im_15 As Currency
    O_16 As Integer
    por_16 As Integer
    im_16 As Currency
    O_17 As Integer
    por_17 As Integer
    im_17 As Currency
    O_18 As Integer
    por_18 As Integer
    im_18 As Currency
    O_19 As Integer
    por_19 As Integer
    im_19 As Currency
    O_20 As Integer
    por_20 As Integer
    im_20 As Currency

 End Type
 Type CAT_MA
    B1 As String * 6
    B2 As String * 32
    B3 As String * 16
    B4 As String * 5
    B5 As String * 5
End Type
Type CAT_AX
    C1 As String * 6
    C2 As String * 32
    C3 As String * 16
    C4 As String * 5
    C5 As String * 5
End Type
Type Clabnx
    Q1 As String * 16
 End Type
 Type da_id
       Emp_Rfc As String * 25
       Emp_Dom As String * 70
       Rep_Legapp As String * 20
       Rep_Legapm As String * 20
       Rep_Legapn As String * 20
       Rep_Rfc As String * 25
       Rep_Curp As String * 25
       suc As String * 4
       cta As String * 12
  End Type
 Public Dat_ide As da_id
 Public CATMAY As CAT_MA, Clbnx As Clabnx
 Public CATAUX As CAT_AX
 Public basico As basini
  Public obras As ob
 Public maestro As ob
 Public empresa As empre, ingresos As Currency
 Public nomina As nom, deducciones As Currency
 Public personal As per, neto As Currency
 Public rgtro As Integer
 Public articulo As art, final As Long
 Public subsidio As subs, impto As Currency
 Public credito As cred, base As Currency
 Public z1 As String, mm(12) As String * 20, cm, dm, z2$, ddm
 Public subdirectorio$, valor$, dd(12) As Integer
 Public Dir_imptos As String
 Public arch_tr As String * 20, tar As Integer, Direc_torio
 Public dir_obras As String, baseanual As Currency, cal_anual As Integer, baseor As Currency
 Public obra(22) As Integer, porcentaje(22) As Currency, Aplicar As Integer
Sub convierte()
obra(1) = maestro.O_1: obra(2) = maestro.O_2: obra(3) = maestro.O_3: obra(4) = maestro.O_4
obra(5) = maestro.O_5: obra(6) = maestro.O_6: obra(7) = maestro.O_7: obra(8) = maestro.O_8
obra(9) = maestro.O_9: obra(10) = maestro.O_10: obra(11) = maestro.O_11: obra(12) = maestro.O_12
obra(13) = maestro.O_13: obra(14) = maestro.O_14: obra(15) = maestro.O_15: obra(16) = maestro.O_16
obra(17) = maestro.O_17: obra(18) = maestro.O_18: obra(19) = maestro.O_19: obra(20) = maestro.O_20
porcentaje(1) = maestro.por_1: porcentaje(2) = maestro.por_2: porcentaje(3) = maestro.por_3: porcentaje(4) = maestro.por_4
porcentaje(5) = maestro.por_5: porcentaje(6) = maestro.por_6: porcentaje(7) = maestro.por_7: porcentaje(8) = maestro.por_8
porcentaje(9) = maestro.por_9: porcentaje(10) = maestro.por_10: porcentaje(11) = maestro.por_11: porcentaje(12) = maestro.por_12
porcentaje(13) = maestro.por_13: porcentaje(14) = maestro.por_14: porcentaje(15) = maestro.por_15: porcentaje(16) = maestro.por_16
porcentaje(17) = maestro.por_17: porcentaje(18) = maestro.por_18: porcentaje(19) = maestro.por_19: porcentaje(20) = maestro.por_20

   
End Sub
Sub grabamaestro()
    maestro.O_1 = obra(1): maestro.O_2 = obra(2): maestro.O_3 = obra(3): maestro.O_4 = obra(4)
    maestro.O_5 = obra(5): maestro.O_6 = obra(6): maestro.O_7 = obra(7): maestro.O_8 = obra(8)
    maestro.O_9 = obra(9): maestro.O_10 = obra(10): maestro.O_11 = obra(11): maestro.O_12 = obra(12)
    maestro.O_13 = obra(13): maestro.O_14 = obra(14): maestro.O_15 = obra(15): maestro.O_16 = obra(16)
    maestro.O_17 = obra(17): maestro.O_18 = obra(18): maestro.O_19 = obra(19): maestro.O_20 = obra(20)
    maestro.por_1 = porcentaje(1): maestro.por_2 = porcentaje(2): maestro.por_3 = porcentaje(3): maestro.por_4 = porcentaje(4)
    maestro.por_5 = porcentaje(5): maestro.por_6 = porcentaje(6): maestro.por_7 = porcentaje(7): maestro.por_8 = porcentaje(8)
    maestro.por_9 = porcentaje(9): maestro.por_10 = porcentaje(10): maestro.por_11 = porcentaje(11): maestro.por_12 = porcentaje(12)
    maestro.por_13 = porcentaje(13): maestro.por_14 = porcentaje(14): maestro.por_15 = porcentaje(15): maestro.por_16 = porcentaje(16)
    maestro.por_17 = porcentaje(17): maestro.por_18 = porcentaje(18): maestro.por_19 = porcentaje(19): maestro.por_20 = porcentaje(20)
End Sub
