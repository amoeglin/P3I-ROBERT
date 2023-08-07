VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTTPROVCOLL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A75FA0356"
Option Explicit

'
' Members
'

'##ModelId=5C8A75FB0068
Public autoMode As Boolean

'##ModelId=5C8A75FB0088
Public m_CDCOMPAGNIE As Variant
'##ModelId=5C8A75FB00A7
Public m_CDAPPLI As Variant
'##ModelId=5C8A75FB00B7
Public m_CDPRODUIT As Variant
'##ModelId=5C8A75FB00D6
Public m_NUCONTRA As Variant
'##ModelId=5C8A75FB00E5
Public m_NUPRET As Variant
'##ModelId=5C8A75FB00F5
Public m_NUDOSSIERPREST As Variant
'##ModelId=5C8A75FB0114
Public m_NUSOUSDOSSIERPREST As Variant
'##ModelId=5C8A75FB0124
Public m_NUCONTRATGESTDELEG As Variant
'##ModelId=5C8A75FB0143
Public m_CDGARAN As Variant
'##ModelId=5C8A75FB0153
Public m_NUREFMVT As Variant
'##ModelId=5C8A75FB0172
Public m_CDTYPMVT As Variant
'##ModelId=5C8A75FB0182
Public m_NUTRAITE As Variant
'##ModelId=5C8A75FB01A1
Public m_CDCATADHESION As Variant
'##ModelId=5C8A75FB01B1
Public m_LBCATADHESION As Variant
'##ModelId=5C8A75FB01D0
Public m_CDOPTION As Variant
'##ModelId=5C8A75FB01DF
Public m_LBOPTIONCON As Variant
'##ModelId=5C8A75FB01FF
Public m_LBSOUSCR As Variant
'##ModelId=5C8A75FB020E
Public m_IDASSUREAGI As Variant
'##ModelId=5C8A75FB023D
Public m_IDASSURE As Variant
'##ModelId=5C8A75FB024D
Public m_LBASSURE As Variant
'##ModelId=5C8A75FB026C
Public m_DTNAISSASS As Variant
'##ModelId=5C8A75FB027C
Public m_CDSEXASSURE As Variant
'##ModelId=5C8A75FB029B
Public m_IDRENTIERAGI As Variant
'##ModelId=5C8A75FB02BA
Public m_IDRENTIER As Variant
'##ModelId=5C8A75FB02CA
Public m_LBRENTIER As Variant
'##ModelId=5C8A75FB02E8
Public m_DTNAISSREN As Variant
'##ModelId=5C8A75FB02F8
Public m_IDCORENTIERAGI As Variant
'##ModelId=5C8A75FB0317
Public m_IDCORENTIER As Variant
'##ModelId=5C8A75FB0336
Public m_LBCORENTIER As Variant
'##ModelId=5C8A75FB0355
Public m_DTNAISSCOR As Variant
'##ModelId=5C8A75FB0375
Public m_DTSURVSIN As Variant
'##ModelId=5C8A75FB0394
Public m_AGESURVSIN As Variant
'##ModelId=5C8A75FB03A4
Public m_DURPOSARR As Variant
'##ModelId=5C8A75FB03C3
Public m_ANCARRTRA As Variant
'##ModelId=5C8A75FB03E2
Public m_CDPERIODICITE As Variant
'##ModelId=5C8A75FC0019
Public m_CDTYPTERME As Variant
'##ModelId=5C8A75FC0039
Public m_DTEFFREN As Variant
'##ModelId=5C8A75FC0058
Public m_DTLIMPROAVR As Variant
'##ModelId=5C8A75FC0077
Public m_DTDERREG As Variant
'##ModelId=5C8A75FC0096
Public m_DTDEBPER As Variant
'##ModelId=5C8A75FC00B6
Public m_DTFINPER As Variant
'##ModelId=5C8A75FC00C5
Public m_CDPROEVA As Variant
'##ModelId=5C8A75FC00E4
Public m_ANNEADHE As Variant
'##ModelId=5C8A75FC0104
Public m_CDSINCON As Variant
'##ModelId=5C8A75FC0123
Public m_CDMISINV As Variant
'##ModelId=5C8A75FC0133
Public m_DTMISINV As Variant
'##ModelId=5C8A75FC0161
Public m_NUETABLI As Variant
'##ModelId=5C8A75FC0181
Public m_TXREVERSION As Variant
'##ModelId=5C8A75FC01A0
Public m_DTCALCULPROV As Variant
'##ModelId=5C8A75FC01B0
Public m_DTTRAITPROV As Variant
'##ModelId=5C8A75FC01CF
Public m_DTCREATI As Variant
'##ModelId=5C8A75FC01FE
Public m_DTSIGBIA As Variant
'##ModelId=5C8A75FC0220
Public m_DTDECSIN As Variant
'##ModelId=5C8A75FC023F
Public m_CDRISQUE As Variant
'##ModelId=5C8A75FC026F
Public m_LBRISQUE As Variant
'##ModelId=5C8A75FC028E
Public m_CDSITUATSIN As Variant
'##ModelId=5C8A75FC029E
Public m_DTSITUATSIN As Variant
'##ModelId=5C8A75FC02D2
Public m_CDPRETRATTSIN As Variant
'##ModelId=5C8A75FC02EA
Public m_CDCTGPRT As Variant
'##ModelId=5C8A75FC030A
Public m_LBCTGPRT As Variant
'##ModelId=5C8A75FC0329
Public m_DTPREECH As Variant
'##ModelId=5C8A75FC0348
Public m_DTDERECH As Variant
'##ModelId=5C8A75FC0367
Public m_CDPERIODICITEECH As Variant
'##ModelId=5C8A75FC0387
Public m_MTECHEANCE1 As Variant
'##ModelId=5C8A75FC03A6
Public m_DTDEBPERECH1 As Variant
'##ModelId=5C8A75FC03C5
Public m_DTFINPERECH1 As Variant
'##ModelId=5C8A75FC03E4
Public m_MTECHEANCE2 As Variant
'##ModelId=5C8A75FD001C
Public m_DTDEBPERECH2 As Variant
'##ModelId=5C8A75FD003B
Public m_DTFINPERECH2 As Variant
'##ModelId=5C8A75FD006A
Public m_MTECHEANCE3 As Variant
'##ModelId=5C8A75FD0089
Public m_DTDEBPERECH3 As Variant
'##ModelId=5C8A75FD00B8
Public m_DTFINPERECH3 As Variant
'##ModelId=5C8A75FD00C8
Public m_CDTYPAMO As Variant
'##ModelId=5C8A75FD00E7
Public m_LBTYPAMO As Variant
'##ModelId=5C8A75FD0106
Public m_TXINVPEC As Variant
'##ModelId=5C8A75FD0145
Public m_DTDEBPERPIP As Variant
'##ModelId=5C8A75FD0164
Public m_DTFINPERPIP As Variant
'##ModelId=5C8A75FD0183
Public m_DTSAISIEPERJUSTIF As Variant
'##ModelId=5C8A75FD01A2
Public m_DTDEBPERJUSTIF As Variant
'##ModelId=5C8A75FD01C2
Public m_DTFINPERJUSTIF As Variant
'##ModelId=5C8A75FD01E1
Public m_DTDEBDERPERRGLTADA As Variant
'##ModelId=5C8A75FD0200
Public m_DTFINDERPERRGLTADA As Variant
'##ModelId=5C8A75FD0217
Public m_DTDERPERRGLTADA As Variant
'##ModelId=5C8A75FD0260
Public m_MTDERPERRGLTADA As Variant
'##ModelId=5C8A75FD027F
Public m_DTDEBDERPERRGLTADC As Variant
'##ModelId=5C8A75FD02AE
Public m_DTFINDERPERRGLTADC As Variant
'##ModelId=5C8A75FD02E0
Public m_DTDERPERRGLTADC As Variant
'##ModelId=5C8A75FD0300
Public m_MTDERPERRGLTADC As Variant
'##ModelId=5C8A75FD031F
Public m_CDSINPREPROV As Variant
'##ModelId=5C8A75FD033E
Public m_MTTOTREGLEICIV As Variant
'##ModelId=5C8A75FD035D
Public m_DTDEBPROV As Variant
'##ModelId=5C8A75FD038C
Public m_DTFINPROV As Variant
'##ModelId=5C8A75FD03AB
Public m_INDBASREV As Variant
'##ModelId=5C8A75FD03CB
Public m_MTPREANN As Variant
'##ModelId=5C8A75FE0002
Public m_MTPREANN2 As Variant
'##ModelId=5C8A75FE0031
Public m_MTCPLEST As Variant
'##ModelId=5C8A75FE0050
Public m_MTCPLREE As Variant
'##ModelId=5C8A75FE006F
Public m_MTPROCAL As Variant
'##ModelId=5C8A75FE009E
Public m_MTPROIMP As Variant
'##ModelId=5C8A75FE00BD
Public m_MTPROVIT As Variant
'##ModelId=5C8A75FE00DD
Public m_MTPROPAS As Variant
'##ModelId=5C8A75FE010C
Public m_TXPROV As Variant
'##ModelId=5C8A75FE012B
Public m_TXPROPASS As Variant
'##ModelId=5C8A75FE015A
Public m_TXTECHN As Variant
'##ModelId=5C8A75FE0179
Public m_TXFRAIS As Variant
'##ModelId=5C8A75FE01A8
Public m_MTCAPCONORI As Variant
'##ModelId=5C8A75FE01C7
Public m_MTPREREV As Variant
'##ModelId=5C8A75FE01E6
Public m_MTPREMAJ As Variant
'##ModelId=5C8A75FE0218
Public m_MTPANCAL As Variant
'##ModelId=5C8A75FE023A
Public m_MTCMPPAN As Variant
'##ModelId=5C8A75FE0259
Public m_MTCOMPAN As Variant
'##ModelId=5C8A75FE0288
Public m_MTCOMCMP As Variant
'##ModelId=5C8A75FE02A7
Public m_MTPRIREG As Variant
'##ModelId=5C8A75FE02D0
Public m_MTPRIRE1 As Variant
'##ModelId=5C8A75FE02FF
Public m_MTPRIRE2 As Variant
'##ModelId=5C8A75FE031E
Public m_CDMONNAIE As Variant
'##ModelId=5C8A75FE034D
Public m_CDPAYS As Variant
'##ModelId=5C8A75FE036C
Public m_CDAPPLISOURCE As Variant
'##ModelId=5C8A75FE038C
Public m_LBCOMLIG As Variant
'##ModelId=5C8A75FE03BB
Public m_CDLOTIMPORT As Variant
'##ModelId=5C8A75FF0001
Public m_NUTRAITP3I As Variant
'##ModelId=5C8A75FF0021
Public m_NUENRP3I As Variant
'##ModelId=5C8A75FF0050
Public m_TYPEPROVISION As Variant
'##ModelId=5C8A75FF007F
Public m_PERALIM As Variant
'##ModelId=5C8A75FF009E
Public m_TXLISSAGE As Variant
'##ModelId=5C8A75FF00CD
Public m_TXPROSANSLISSAGE As Variant
'##ModelId=5C8A75FF00EC
Public m_CDCATINV As Variant
'##ModelId=5C8A75FF011B
Public m_LBCATINV As Variant
'##ModelId=5C8A75FF013A
Public m_CDCONTENTIEUX As Variant
'##ModelId=5C8A75FF0169
Public m_NUSINISTRE As Variant
'##ModelId=5C8A75FF0188
Public m_CDCHOIXPREST As Variant
'##ModelId=5C8A75FF01B7
Public m_LBCHOIXPREST As Variant
'##ModelId=5C8A75FF01E6
Public m_MTCAPCON As Variant
'##ModelId=5C8A75FF0227
Public m_MTCAPSSRISQ As Variant

'##ModelId=5C8A75FF0256
Public m_FLAMORTISSABLE As Variant
'##ModelId=5C8A75FF0275
Public m_TXAMORT As Variant
'##ModelId=5C8A75FF02A4
Public m_DTLIMPROAPR As Variant
'##ModelId=5C8A75FF02CD
Public m_MTPROIMPAVR As Variant
'##ModelId=5C8A75FF02FB
Public m_MTPROIMPAPR As Variant
'##ModelId=5C8A75FF032A
Public m_MTINDEMRES As Variant
'##ModelId=5C8A75FF034A
Public m_MTPROCALAVR As Variant
'##ModelId=5C8A75FF0378
Public m_MTPROCALAPR As Variant
'##ModelId=5C8A75FF03A7
Public m_MTCOUTREFPROCAL As Variant
'##ModelId=5C8A75FF03D6
Public m_MTAMORTCAL As Variant
'##ModelId=5C8A7600000D
Public m_MTRESTEAMORCAL As Variant
'##ModelId=5C8A7600003C
Public m_MTPROPASAVR As Variant
'##ModelId=5C8A7600006B
Public m_MTPROPASAPR As Variant
'##ModelId=5C8A7600009A
Public m_MTCOUTREFPROPAS As Variant
'##ModelId=5C8A760000B9
Public m_MTAMORTPAS As Variant
'##ModelId=5C8A760000E8
Public m_MTRESTEAMORPAS As Variant

'##ModelId=5C8A76000117
Public m_AGELIMINV As Variant
'##ModelId=5C8A76000146
Public m_AGELIMINC As Variant

'##ModelId=5C8A76000165
Public m_COMMENTAIRE As Variant
'
'


'
' Constructors
'

'##ModelId=5C8A76000194
Public Sub Copy_OLD(ByRef src As clsTTPROVCOLL)

  m_CDCOMPAGNIE = src.m_CDCOMPAGNIE
  m_CDAPPLI = src.m_CDAPPLI
  m_CDPRODUIT = src.m_CDPRODUIT
  m_NUCONTRA = src.m_NUCONTRA
  m_NUPRET = src.m_NUPRET
  m_NUDOSSIERPREST = src.m_NUDOSSIERPREST
  m_NUSOUSDOSSIERPREST = src.m_NUSOUSDOSSIERPREST
  m_NUCONTRATGESTDELEG = src.m_NUCONTRATGESTDELEG
  m_CDGARAN = src.m_CDGARAN
  m_NUREFMVT = src.m_NUREFMVT
  m_CDTYPMVT = src.m_CDTYPMVT
  m_NUTRAITE = src.m_NUTRAITE
  m_CDCATADHESION = src.m_CDCATADHESION
  m_LBCATADHESION = src.m_LBCATADHESION
  m_CDOPTION = src.m_CDOPTION
  m_LBOPTIONCON = src.m_LBOPTIONCON
  m_LBSOUSCR = src.m_LBSOUSCR
  m_IDASSUREAGI = src.m_IDASSUREAGI
  m_IDASSURE = src.m_IDASSURE
  m_LBASSURE = src.m_LBASSURE
  m_DTNAISSASS = src.m_DTNAISSASS
  m_CDSEXASSURE = src.m_CDSEXASSURE
  m_IDRENTIERAGI = src.m_IDRENTIERAGI
  m_IDRENTIER = src.m_IDRENTIER
  m_LBRENTIER = src.m_LBRENTIER
  m_DTNAISSREN = src.m_DTNAISSREN
  m_IDCORENTIERAGI = src.m_IDCORENTIERAGI
  m_IDCORENTIER = src.m_IDCORENTIER
  m_LBCORENTIER = src.m_LBCORENTIER
  m_DTNAISSCOR = src.m_DTNAISSCOR
  m_DTSURVSIN = src.m_DTSURVSIN
  m_AGESURVSIN = src.m_AGESURVSIN
  m_DURPOSARR = src.m_DURPOSARR
  m_ANCARRTRA = src.m_ANCARRTRA
  m_CDPERIODICITE = src.m_CDPERIODICITE
  m_CDTYPTERME = src.m_CDTYPTERME
  m_DTEFFREN = src.m_DTEFFREN
  m_DTLIMPROAVR = src.m_DTLIMPROAVR
  m_DTDERREG = src.m_DTDERREG
  m_DTDEBPER = src.m_DTDEBPER
  m_DTFINPER = src.m_DTFINPER
  m_CDPROEVA = src.m_CDPROEVA
  m_ANNEADHE = src.m_ANNEADHE
  m_CDSINCON = src.m_CDSINCON
  m_CDMISINV = src.m_CDMISINV
  m_DTMISINV = src.m_DTMISINV
  m_NUETABLI = src.m_NUETABLI
  m_TXREVERSION = src.m_TXREVERSION
  m_DTCALCULPROV = src.m_DTCALCULPROV
  m_DTTRAITPROV = src.m_DTTRAITPROV
  m_DTCREATI = src.m_DTCREATI
  m_DTSIGBIA = src.m_DTSIGBIA
  m_DTDECSIN = src.m_DTDECSIN
  m_CDRISQUE = src.m_CDRISQUE
  m_LBRISQUE = src.m_LBRISQUE
  m_CDSITUATSIN = src.m_CDSITUATSIN
  m_DTSITUATSIN = src.m_DTSITUATSIN
  m_CDPRETRATTSIN = src.m_CDPRETRATTSIN
  m_CDCTGPRT = src.m_CDCTGPRT
  m_LBCTGPRT = src.m_LBCTGPRT
  m_DTPREECH = src.m_DTPREECH
  m_DTDERECH = src.m_DTDERECH
  m_CDPERIODICITEECH = src.m_CDPERIODICITEECH
  m_MTECHEANCE1 = src.m_MTECHEANCE1
  m_DTDEBPERECH1 = src.m_DTDEBPERECH1
  m_DTFINPERECH1 = src.m_DTFINPERECH1
  m_MTECHEANCE2 = src.m_MTECHEANCE2
  m_DTDEBPERECH2 = src.m_DTDEBPERECH2
  m_DTFINPERECH2 = src.m_DTFINPERECH2
  m_MTECHEANCE3 = src.m_MTECHEANCE3
  m_DTDEBPERECH3 = src.m_DTDEBPERECH3
  m_DTFINPERECH3 = src.m_DTFINPERECH3
  m_CDTYPAMO = src.m_CDTYPAMO
  m_LBTYPAMO = src.m_LBTYPAMO
  m_TXINVPEC = src.m_TXINVPEC
  m_DTDEBPERPIP = src.m_DTDEBPERPIP
  m_DTFINPERPIP = src.m_DTFINPERPIP
  m_DTSAISIEPERJUSTIF = src.m_DTSAISIEPERJUSTIF
  m_DTDEBPERJUSTIF = src.m_DTDEBPERJUSTIF
  m_DTFINPERJUSTIF = src.m_DTFINPERJUSTIF
  m_DTDEBDERPERRGLTADA = src.m_DTDEBDERPERRGLTADA
  m_DTFINDERPERRGLTADA = src.m_DTFINDERPERRGLTADA
  m_DTDERPERRGLTADA = src.m_DTDERPERRGLTADA
  m_MTDERPERRGLTADA = src.m_MTDERPERRGLTADA
  m_DTDEBDERPERRGLTADC = src.m_DTDEBDERPERRGLTADC
  m_DTFINDERPERRGLTADC = src.m_DTFINDERPERRGLTADC
  m_DTDERPERRGLTADC = src.m_DTDERPERRGLTADC
  m_MTDERPERRGLTADC = src.m_MTDERPERRGLTADC
  m_CDSINPREPROV = src.m_CDSINPREPROV
  m_MTTOTREGLEICIV = src.m_MTTOTREGLEICIV
  m_DTDEBPROV = src.m_DTDEBPROV
  m_DTFINPROV = src.m_DTFINPROV
  m_INDBASREV = src.m_INDBASREV
  m_MTPREANN = src.m_MTPREANN
  m_MTPREANN2 = src.m_MTPREANN2
  m_MTCPLEST = src.m_MTCPLEST
  m_MTCPLREE = src.m_MTCPLREE
  m_MTPROCAL = src.m_MTPROCAL
  m_MTPROIMP = src.m_MTPROIMP
  m_MTPROVIT = src.m_MTPROVIT
  m_MTPROPAS = src.m_MTPROPAS
  m_TXPROV = src.m_TXPROV
  m_TXPROPASS = src.m_TXPROPASS
  m_TXTECHN = src.m_TXTECHN
  m_TXFRAIS = src.m_TXFRAIS
  m_MTCAPCONORI = src.m_MTCAPCONORI
  m_MTPREREV = src.m_MTPREREV
  m_MTPREMAJ = src.m_MTPREMAJ
  m_MTPANCAL = src.m_MTPANCAL
  m_MTCMPPAN = src.m_MTCMPPAN
  m_MTCOMPAN = src.m_MTCOMPAN
  m_MTCOMCMP = src.m_MTCOMCMP
  m_MTPRIREG = src.m_MTPRIREG
  m_MTPRIRE1 = src.m_MTPRIRE1
  m_MTPRIRE2 = src.m_MTPRIRE2
  m_CDMONNAIE = src.m_CDMONNAIE
  m_CDPAYS = src.m_CDPAYS
  m_CDAPPLISOURCE = src.m_CDAPPLISOURCE
  m_LBCOMLIG = src.m_LBCOMLIG
  m_CDLOTIMPORT = src.m_CDLOTIMPORT
  m_NUTRAITP3I = src.m_NUTRAITP3I
  m_NUENRP3I = src.m_NUENRP3I
  m_TYPEPROVISION = src.m_TYPEPROVISION
  m_PERALIM = src.m_PERALIM
  m_TXLISSAGE = src.m_TXLISSAGE
  m_TXPROSANSLISSAGE = src.m_TXPROSANSLISSAGE
  
  m_CDCATINV = src.m_CDCATINV
  m_LBCATINV = src.m_LBCATINV
  m_CDCONTENTIEUX = src.m_CDCONTENTIEUX
  m_NUSINISTRE = src.m_NUSINISTRE
  m_CDCHOIXPREST = src.m_CDCHOIXPREST
  m_LBCHOIXPREST = src.m_LBCHOIXPREST
  m_MTCAPCON = src.m_MTCAPCON
  
  m_MTCAPSSRISQ = src.m_MTCAPSSRISQ
  
  m_FLAMORTISSABLE = src.m_FLAMORTISSABLE
  m_TXAMORT = src.m_TXAMORT
  m_DTLIMPROAPR = src.m_DTLIMPROAPR
  m_MTPROIMPAVR = src.m_MTPROIMPAVR
  m_MTPROIMPAPR = src.m_MTPROIMPAPR
  m_MTINDEMRES = src.m_MTINDEMRES
  m_MTPROCALAVR = src.m_MTPROCALAVR
  m_MTPROCALAPR = src.m_MTPROCALAPR
  m_MTCOUTREFPROCAL = src.m_MTCOUTREFPROCAL
  m_MTAMORTCAL = src.m_MTAMORTCAL
  m_MTRESTEAMORCAL = src.m_MTRESTEAMORCAL
  m_MTPROPASAVR = src.m_MTPROPASAVR
  m_MTPROPASAPR = src.m_MTPROPASAPR
  m_MTCOUTREFPROPAS = src.m_MTCOUTREFPROPAS
  m_MTAMORTPAS = src.m_MTAMORTPAS
  m_MTRESTEAMORPAS = src.m_MTRESTEAMORPAS
  
  m_AGELIMINC = src.m_AGELIMINC
  m_AGELIMINV = src.m_AGELIMINV
  
End Sub


'##ModelId=5C8A760001E2
Public Sub Copy(ByRef src As clsTTPROVCOLL)

'Replace$(theData.fields("CDAPPLI"), ",", ".")
  m_CDCOMPAGNIE = Replace$(src.m_CDCOMPAGNIE, ",", ".")
  m_CDAPPLI = Replace$(src.m_CDAPPLI, ",", ".")
  m_CDPRODUIT = Replace$(src.m_CDPRODUIT, ",", ".")
  m_NUCONTRA = Replace$(src.m_NUCONTRA, ",", ".")
  m_NUPRET = Replace$(src.m_NUPRET, ",", ".")
  m_NUDOSSIERPREST = Replace$(src.m_NUDOSSIERPREST, ",", ".")
  m_NUSOUSDOSSIERPREST = Replace$(src.m_NUSOUSDOSSIERPREST, ",", ".")
  m_NUCONTRATGESTDELEG = Replace$(src.m_NUCONTRATGESTDELEG, ",", ".")
  m_CDGARAN = Replace$(src.m_CDGARAN, ",", ".")
  m_NUREFMVT = Replace$(src.m_NUREFMVT, ",", ".")
  m_CDTYPMVT = Replace$(src.m_CDTYPMVT, ",", ".")
  m_NUTRAITE = Replace$(src.m_NUTRAITE, ",", ".")
  m_CDCATADHESION = Replace$(src.m_CDCATADHESION, ",", ".")
  m_LBCATADHESION = Replace$(src.m_LBCATADHESION, ",", ".")
  m_CDOPTION = Replace$(src.m_CDOPTION, ",", ".")
  m_LBOPTIONCON = Replace$(src.m_LBOPTIONCON, ",", ".")
  m_LBSOUSCR = Replace$(src.m_LBSOUSCR, ",", ".")
  m_IDASSUREAGI = Replace$(src.m_IDASSUREAGI, ",", ".")
  m_IDASSURE = Replace$(src.m_IDASSURE, ",", ".")
  m_LBASSURE = Replace$(src.m_LBASSURE, ",", ".")
  m_DTNAISSASS = Replace$(src.m_DTNAISSASS, ",", ".")
  m_CDSEXASSURE = Replace$(src.m_CDSEXASSURE, ",", ".")
  m_IDRENTIERAGI = Replace$(src.m_IDRENTIERAGI, ",", ".")
  m_IDRENTIER = Replace$(src.m_IDRENTIER, ",", ".")
  m_LBRENTIER = Replace$(src.m_LBRENTIER, ",", ".")
  m_DTNAISSREN = Replace$(src.m_DTNAISSREN, ",", ".")
  m_IDCORENTIERAGI = Replace$(src.m_IDCORENTIERAGI, ",", ".")
  m_IDCORENTIER = Replace$(src.m_IDCORENTIER, ",", ".")
  m_LBCORENTIER = Replace$(src.m_LBCORENTIER, ",", ".")
  m_DTNAISSCOR = Replace$(src.m_DTNAISSCOR, ",", ".")
  m_DTSURVSIN = Replace$(src.m_DTSURVSIN, ",", ".")
  m_AGESURVSIN = Replace$(src.m_AGESURVSIN, ",", ".")
  m_DURPOSARR = Replace$(src.m_DURPOSARR, ",", ".")
  m_ANCARRTRA = Replace$(src.m_ANCARRTRA, ",", ".")
  m_CDPERIODICITE = Replace$(src.m_CDPERIODICITE, ",", ".")
  m_CDTYPTERME = Replace$(src.m_CDTYPTERME, ",", ".")
  m_DTEFFREN = Replace$(src.m_DTEFFREN, ",", ".")
  m_DTLIMPROAVR = Replace$(src.m_DTLIMPROAVR, ",", ".")
  m_DTDERREG = Replace$(src.m_DTDERREG, ",", ".")
  m_DTDEBPER = Replace$(src.m_DTDEBPER, ",", ".")
  m_DTFINPER = Replace$(src.m_DTFINPER, ",", ".")
  m_CDPROEVA = Replace$(src.m_CDPROEVA, ",", ".")
  m_ANNEADHE = Replace$(src.m_ANNEADHE, ",", ".")
  m_CDSINCON = Replace$(src.m_CDSINCON, ",", ".")
  m_CDMISINV = Replace$(src.m_CDMISINV, ",", ".")
  m_DTMISINV = Replace$(src.m_DTMISINV, ",", ".")
  m_NUETABLI = Replace$(src.m_NUETABLI, ",", ".")
  m_TXREVERSION = Replace$(src.m_TXREVERSION, ",", ".")
  m_DTCALCULPROV = Replace$(src.m_DTCALCULPROV, ",", ".")
  m_DTTRAITPROV = Replace$(src.m_DTTRAITPROV, ",", ".")
  m_DTCREATI = Replace$(src.m_DTCREATI, ",", ".")
  m_DTSIGBIA = Replace$(src.m_DTSIGBIA, ",", ".")
  m_DTDECSIN = Replace$(src.m_DTDECSIN, ",", ".")
  m_CDRISQUE = Replace$(src.m_CDRISQUE, ",", ".")
  m_LBRISQUE = Replace$(src.m_LBRISQUE, ",", ".")
  m_CDSITUATSIN = Replace$(src.m_CDSITUATSIN, ",", ".")
  m_DTSITUATSIN = Replace$(src.m_DTSITUATSIN, ",", ".")
  m_CDPRETRATTSIN = Replace$(src.m_CDPRETRATTSIN, ",", ".")
  m_CDCTGPRT = Replace$(src.m_CDCTGPRT, ",", ".")
  m_LBCTGPRT = Replace$(src.m_LBCTGPRT, ",", ".")
  m_DTPREECH = Replace$(src.m_DTPREECH, ",", ".")
  m_DTDERECH = Replace$(src.m_DTDERECH, ",", ".")
  m_CDPERIODICITEECH = Replace$(src.m_CDPERIODICITEECH, ",", ".")
  m_MTECHEANCE1 = Replace$(src.m_MTECHEANCE1, ",", ".")
  m_DTDEBPERECH1 = Replace$(src.m_DTDEBPERECH1, ",", ".")
  m_DTFINPERECH1 = Replace$(src.m_DTFINPERECH1, ",", ".")
  m_MTECHEANCE2 = Replace$(src.m_MTECHEANCE2, ",", ".")
  m_DTDEBPERECH2 = Replace$(src.m_DTDEBPERECH2, ",", ".")
  m_DTFINPERECH2 = Replace$(src.m_DTFINPERECH2, ",", ".")
  m_MTECHEANCE3 = Replace$(src.m_MTECHEANCE3, ",", ".")
  m_DTDEBPERECH3 = Replace$(src.m_DTDEBPERECH3, ",", ".")
  m_DTFINPERECH3 = Replace$(src.m_DTFINPERECH3, ",", ".")
  m_CDTYPAMO = Replace$(src.m_CDTYPAMO, ",", ".")
  m_LBTYPAMO = Replace$(src.m_LBTYPAMO, ",", ".")
  m_TXINVPEC = Replace$(src.m_TXINVPEC, ",", ".")
  m_DTDEBPERPIP = Replace$(src.m_DTDEBPERPIP, ",", ".")
  m_DTFINPERPIP = Replace$(src.m_DTFINPERPIP, ",", ".")
  m_DTSAISIEPERJUSTIF = Replace$(src.m_DTSAISIEPERJUSTIF, ",", ".")
  m_DTDEBPERJUSTIF = Replace$(src.m_DTDEBPERJUSTIF, ",", ".")
  m_DTFINPERJUSTIF = Replace$(src.m_DTFINPERJUSTIF, ",", ".")
  m_DTDEBDERPERRGLTADA = Replace$(src.m_DTDEBDERPERRGLTADA, ",", ".")
  m_DTFINDERPERRGLTADA = Replace$(src.m_DTFINDERPERRGLTADA, ",", ".")
  m_DTDERPERRGLTADA = Replace$(src.m_DTDERPERRGLTADA, ",", ".")
  m_MTDERPERRGLTADA = Replace$(src.m_MTDERPERRGLTADA, ",", ".")
  m_DTDEBDERPERRGLTADC = Replace$(src.m_DTDEBDERPERRGLTADC, ",", ".")
  m_DTFINDERPERRGLTADC = Replace$(src.m_DTFINDERPERRGLTADC, ",", ".")
  m_DTDERPERRGLTADC = Replace$(src.m_DTDERPERRGLTADC, ",", ".")
  m_MTDERPERRGLTADC = Replace$(src.m_MTDERPERRGLTADC, ",", ".")
  m_CDSINPREPROV = Replace$(src.m_CDSINPREPROV, ",", ".")
  m_MTTOTREGLEICIV = Replace$(src.m_MTTOTREGLEICIV, ",", ".")
  m_DTDEBPROV = Replace$(src.m_DTDEBPROV, ",", ".")
  m_DTFINPROV = Replace$(src.m_DTFINPROV, ",", ".")
  m_INDBASREV = Replace$(src.m_INDBASREV, ",", ".")
  m_MTPREANN = Replace$(src.m_MTPREANN, ",", ".")
  m_MTPREANN2 = Replace$(src.m_MTPREANN2, ",", ".")
  m_MTCPLEST = Replace$(src.m_MTCPLEST, ",", ".")
  m_MTCPLREE = Replace$(src.m_MTCPLREE, ",", ".")
  m_MTPROCAL = Replace$(src.m_MTPROCAL, ",", ".")
  m_MTPROIMP = Replace$(src.m_MTPROIMP, ",", ".")
  m_MTPROVIT = Replace$(src.m_MTPROVIT, ",", ".")
  m_MTPROPAS = Replace$(src.m_MTPROPAS, ",", ".")
  m_TXPROV = Replace$(src.m_TXPROV, ",", ".")
  m_TXPROPASS = Replace$(src.m_TXPROPASS, ",", ".")
  m_TXTECHN = Replace$(src.m_TXTECHN, ",", ".")
  m_TXFRAIS = Replace$(src.m_TXFRAIS, ",", ".")
  m_MTCAPCONORI = Replace$(src.m_MTCAPCONORI, ",", ".")
  m_MTPREREV = Replace$(src.m_MTPREREV, ",", ".")
  m_MTPREMAJ = Replace$(src.m_MTPREMAJ, ",", ".")
  m_MTPANCAL = Replace$(src.m_MTPANCAL, ",", ".")
  m_MTCMPPAN = Replace$(src.m_MTCMPPAN, ",", ".")
  m_MTCOMPAN = Replace$(src.m_MTCOMPAN, ",", ".")
  m_MTCOMCMP = Replace$(src.m_MTCOMCMP, ",", ".")
  m_MTPRIREG = Replace$(src.m_MTPRIREG, ",", ".")
  m_MTPRIRE1 = Replace$(src.m_MTPRIRE1, ",", ".")
  m_MTPRIRE2 = Replace$(src.m_MTPRIRE2, ",", ".")
  m_CDMONNAIE = Replace$(src.m_CDMONNAIE, ",", ".")
  m_CDPAYS = Replace$(src.m_CDPAYS, ",", ".")
  m_CDAPPLISOURCE = Replace$(src.m_CDAPPLISOURCE, ",", ".")
  m_LBCOMLIG = Replace$(src.m_LBCOMLIG, ",", ".")
  m_CDLOTIMPORT = Replace$(src.m_CDLOTIMPORT, ",", ".")
  m_NUTRAITP3I = Replace$(src.m_NUTRAITP3I, ",", ".")
  m_NUENRP3I = Replace$(src.m_NUENRP3I, ",", ".")
  m_TYPEPROVISION = Replace$(src.m_TYPEPROVISION, ",", ".")
  m_PERALIM = Replace$(src.m_PERALIM, ",", ".")
  m_TXLISSAGE = Replace$(src.m_TXLISSAGE, ",", ".")
  m_TXPROSANSLISSAGE = Replace$(src.m_TXPROSANSLISSAGE, ",", ".")
  
  m_CDCATINV = Replace$(src.m_CDCATINV, ",", ".")
  m_LBCATINV = Replace$(src.m_LBCATINV, ",", ".")
  m_CDCONTENTIEUX = Replace$(src.m_CDCONTENTIEUX, ",", ".")
  m_NUSINISTRE = Replace$(src.m_NUSINISTRE, ",", ".")
  m_CDCHOIXPREST = Replace$(src.m_CDCHOIXPREST, ",", ".")
  m_LBCHOIXPREST = Replace$(src.m_LBCHOIXPREST, ",", ".")
  m_MTCAPCON = Replace$(src.m_MTCAPCON, ",", ".")
  
  m_MTCAPSSRISQ = Replace$(src.m_MTCAPSSRISQ, ",", ".")
  
  m_FLAMORTISSABLE = Replace$(src.m_FLAMORTISSABLE, ",", ".")
  m_TXAMORT = Replace$(src.m_TXAMORT, ",", ".")
  m_DTLIMPROAPR = Replace$(src.m_DTLIMPROAPR, ",", ".")
  m_MTPROIMPAVR = Replace$(src.m_MTPROIMPAVR, ",", ".")
  m_MTPROIMPAPR = Replace$(src.m_MTPROIMPAPR, ",", ".")
  m_MTINDEMRES = Replace$(src.m_MTINDEMRES, ",", ".")
  m_MTPROCALAVR = Replace$(src.m_MTPROCALAVR, ",", ".")
  m_MTPROCALAPR = Replace$(src.m_MTPROCALAPR, ",", ".")
  m_MTCOUTREFPROCAL = Replace$(src.m_MTCOUTREFPROCAL, ",", ".")
  m_MTAMORTCAL = Replace$(src.m_MTAMORTCAL, ",", ".")
  m_MTRESTEAMORCAL = Replace$(src.m_MTRESTEAMORCAL, ",", ".")
  m_MTPROPASAVR = Replace$(src.m_MTPROPASAVR, ",", ".")
  m_MTPROPASAPR = Replace$(src.m_MTPROPASAPR, ",", ".")
  m_MTCOUTREFPROPAS = Replace$(src.m_MTCOUTREFPROPAS, ",", ".")
  m_MTAMORTPAS = Replace$(src.m_MTAMORTPAS, ",", ".")
  m_MTRESTEAMORPAS = Replace$(src.m_MTRESTEAMORPAS, ",", ".")
  
  m_AGELIMINC = Replace$(src.m_AGELIMINC, ",", ".")
  m_AGELIMINV = Replace$(src.m_AGELIMINV, ",", ".")
  m_COMMENTAIRE = Replace$(src.m_COMMENTAIRE, ";", ",")
  
End Sub



'---------------------------------------------------
'##ModelId=5C8A76000217
Public Function Load_OLD(DB As DataAccess, theData As ADODB.Recordset, bFromP3IPROVCOL As Boolean) As Boolean

   'Dim theData As ADODB.Recordset
   Dim ret As Boolean
   ret = True

   On Error GoTo err_Load

   If theData.EOF = False Then

      If IsNull(theData.fields("CDCOMPAGNIE")) Then
         m_CDCOMPAGNIE = 0
      Else
         m_CDCOMPAGNIE = theData.fields("CDCOMPAGNIE")
      End If

      If IsNull(theData.fields("CDAPPLI")) Then
         m_CDAPPLI = 0
      Else
         m_CDAPPLI = theData.fields("CDAPPLI")
      End If

      If IsNull(theData.fields("CDPRODUIT")) Then
         m_CDPRODUIT = ""
      Else
         m_CDPRODUIT = theData.fields("CDPRODUIT")
      End If

      If IsNull(theData.fields("NUCONTRA")) Then
         m_NUCONTRA = ""
      Else
         m_NUCONTRA = theData.fields("NUCONTRA")
      End If

      If IsNull(theData.fields("NUPRET")) Then
         m_NUPRET = 0
      Else
         m_NUPRET = theData.fields("NUPRET")
      End If

      If IsNull(theData.fields("NUDOSSIERPREST")) Then
         m_NUDOSSIERPREST = ""
      Else
         m_NUDOSSIERPREST = theData.fields("NUDOSSIERPREST")
      End If

      If IsNull(theData.fields("NUSOUSDOSSIERPREST")) Then
         m_NUSOUSDOSSIERPREST = ""
      Else
         m_NUSOUSDOSSIERPREST = theData.fields("NUSOUSDOSSIERPREST")
      End If

      If IsNull(theData.fields("NUCONTRATGESTDELEG")) Then
         m_NUCONTRATGESTDELEG = ""
      Else
         m_NUCONTRATGESTDELEG = theData.fields("NUCONTRATGESTDELEG")
      End If

      If IsNull(theData.fields("CDGARAN")) Then
         m_CDGARAN = 0
      Else
         m_CDGARAN = theData.fields("CDGARAN")
      End If

      If IsNull(theData.fields("NUREFMVT")) Then
         m_NUREFMVT = 0
      Else
         m_NUREFMVT = theData.fields("NUREFMVT")
      End If

      If IsNull(theData.fields("CDTYPMVT")) Then
         m_CDTYPMVT = 0
      Else
         m_CDTYPMVT = theData.fields("CDTYPMVT")
      End If

      If IsNull(theData.fields("NUTRAITE")) Then
         m_NUTRAITE = 0
      Else
         m_NUTRAITE = theData.fields("NUTRAITE")
      End If

      If IsNull(theData.fields("CDCATADHESION")) Then
         m_CDCATADHESION = 0
      Else
         m_CDCATADHESION = theData.fields("CDCATADHESION")
      End If

      If IsNull(theData.fields("LBCATADHESION")) Then
         m_LBCATADHESION = ""
      Else
         m_LBCATADHESION = theData.fields("LBCATADHESION")
      End If

      If IsNull(theData.fields("CDOPTION")) Then
         m_CDOPTION = ""
      Else
         m_CDOPTION = theData.fields("CDOPTION")
      End If

      If IsNull(theData.fields("LBOPTIONCON")) Then
         m_LBOPTIONCON = ""
      Else
         m_LBOPTIONCON = theData.fields("LBOPTIONCON")
      End If

      If IsNull(theData.fields("LBSOUSCR")) Then
         m_LBSOUSCR = ""
      Else
         m_LBSOUSCR = theData.fields("LBSOUSCR")
      End If

      If IsNull(theData.fields("IDASSUREAGI")) Then
         m_IDASSUREAGI = ""
      Else
         m_IDASSUREAGI = theData.fields("IDASSUREAGI")
      End If

      If IsNull(theData.fields("IDASSURE")) Then
         m_IDASSURE = 0
      Else
         m_IDASSURE = theData.fields("IDASSURE")
      End If

      If IsNull(theData.fields("LBASSURE")) Then
         m_LBASSURE = ""
      Else
         m_LBASSURE = theData.fields("LBASSURE")
      End If

      If IsNull(theData.fields("DTNAISSASS")) Then
         m_DTNAISSASS = 0
      Else
         m_DTNAISSASS = theData.fields("DTNAISSASS")
      End If

      If IsNull(theData.fields("CDSEXASSURE")) Then
         m_CDSEXASSURE = ""
      Else
         m_CDSEXASSURE = theData.fields("CDSEXASSURE")
      End If

      If IsNull(theData.fields("IDRENTIERAGI")) Then
         m_IDRENTIERAGI = ""
      Else
         m_IDRENTIERAGI = theData.fields("IDRENTIERAGI")
      End If

      If IsNull(theData.fields("IDRENTIER")) Then
         m_IDRENTIER = 0
      Else
         m_IDRENTIER = theData.fields("IDRENTIER")
      End If

      If IsNull(theData.fields("LBRENTIER")) Then
         m_LBRENTIER = ""
      Else
         m_LBRENTIER = theData.fields("LBRENTIER")
      End If

      If IsNull(theData.fields("DTNAISSREN")) Then
         m_DTNAISSREN = 0
      Else
         m_DTNAISSREN = theData.fields("DTNAISSREN")
      End If

      If IsNull(theData.fields("IDCORENTIERAGI")) Then
         m_IDCORENTIERAGI = ""
      Else
         m_IDCORENTIERAGI = theData.fields("IDCORENTIERAGI")
      End If

      If IsNull(theData.fields("IDCORENTIER")) Then
         m_IDCORENTIER = 0
      Else
         m_IDCORENTIER = theData.fields("IDCORENTIER")
      End If

      If IsNull(theData.fields("LBCORENTIER")) Then
         m_LBCORENTIER = ""
      Else
         m_LBCORENTIER = theData.fields("LBCORENTIER")
      End If

      If IsNull(theData.fields("DTNAISSCOR")) Then
         m_DTNAISSCOR = 0
      Else
         m_DTNAISSCOR = theData.fields("DTNAISSCOR")
      End If

      If IsNull(theData.fields("DTSURVSIN")) Then
         m_DTSURVSIN = 0
      Else
         m_DTSURVSIN = theData.fields("DTSURVSIN")
      End If

      If IsNull(theData.fields("AGESURVSIN")) Then
         m_AGESURVSIN = 0
      Else
         m_AGESURVSIN = theData.fields("AGESURVSIN")
      End If

      If IsNull(theData.fields("DURPOSARR")) Then
         m_DURPOSARR = 0
      Else
         m_DURPOSARR = theData.fields("DURPOSARR")
      End If

      If IsNull(theData.fields("ANCARRTRA")) Then
         m_ANCARRTRA = 0
      Else
         m_ANCARRTRA = theData.fields("ANCARRTRA")
      End If

      If IsNull(theData.fields("CDPERIODICITE")) Then
         m_CDPERIODICITE = ""
      Else
         m_CDPERIODICITE = theData.fields("CDPERIODICITE")
      End If

      If IsNull(theData.fields("CDTYPTERME")) Then
         m_CDTYPTERME = 0
      Else
         m_CDTYPTERME = theData.fields("CDTYPTERME")
      End If

      If IsNull(theData.fields("DTEFFREN")) Then
         m_DTEFFREN = 0
      Else
         m_DTEFFREN = theData.fields("DTEFFREN")
      End If

      If IsNull(theData.fields("DTLIMPROAVR")) Then
         m_DTLIMPROAVR = 0
      Else
         m_DTLIMPROAVR = theData.fields("DTLIMPROAVR")
      End If

      If IsNull(theData.fields("DTDERREG")) Then
         m_DTDERREG = 0
      Else
         m_DTDERREG = theData.fields("DTDERREG")
      End If

      If IsNull(theData.fields("DTDEBPER")) Then
         m_DTDEBPER = 0
      Else
         m_DTDEBPER = theData.fields("DTDEBPER")
      End If

      If IsNull(theData.fields("DTFINPER")) Then
         m_DTFINPER = 0
      Else
         m_DTFINPER = theData.fields("DTFINPER")
      End If

      If IsNull(theData.fields("CDPROEVA")) Then
         m_CDPROEVA = ""
      Else
         m_CDPROEVA = theData.fields("CDPROEVA")
      End If

      If IsNull(theData.fields("ANNEADHE")) Then
         m_ANNEADHE = 0
      Else
         m_ANNEADHE = theData.fields("ANNEADHE")
      End If

      If IsNull(theData.fields("CDSINCON")) Then
         m_CDSINCON = ""
      Else
         m_CDSINCON = theData.fields("CDSINCON")
      End If

      If IsNull(theData.fields("CDMISINV")) Then
         m_CDMISINV = ""
      Else
         m_CDMISINV = theData.fields("CDMISINV")
      End If

      If IsNull(theData.fields("DTMISINV")) Then
         m_DTMISINV = 0
      Else
         m_DTMISINV = theData.fields("DTMISINV")
      End If

      If IsNull(theData.fields("NUETABLI")) Then
         m_NUETABLI = 0
      Else
         m_NUETABLI = theData.fields("NUETABLI")
      End If

      If IsNull(theData.fields("TXREVERSION")) Then
         m_TXREVERSION = 0
      Else
         m_TXREVERSION = theData.fields("TXREVERSION")
      End If

      If IsNull(theData.fields("DTCALCULPROV")) Then
         m_DTCALCULPROV = 0
      Else
         m_DTCALCULPROV = theData.fields("DTCALCULPROV")
      End If

      If IsNull(theData.fields("DTTRAITPROV")) Then
         m_DTTRAITPROV = 0
      Else
         m_DTTRAITPROV = theData.fields("DTTRAITPROV")
      End If

      If IsNull(theData.fields("DTCREATI")) Then
         m_DTCREATI = 0
      Else
         m_DTCREATI = theData.fields("DTCREATI")
      End If

      If IsNull(theData.fields("DTSIGBIA")) Then
         m_DTSIGBIA = 0
      Else
         m_DTSIGBIA = theData.fields("DTSIGBIA")
      End If

      If IsNull(theData.fields("DTDECSIN")) Then
         m_DTDECSIN = 0
      Else
         m_DTDECSIN = theData.fields("DTDECSIN")
      End If

      If IsNull(theData.fields("CDRISQUE")) Then
         m_CDRISQUE = ""
      Else
         m_CDRISQUE = theData.fields("CDRISQUE")
      End If

      If IsNull(theData.fields("LBRISQUE")) Then
         m_LBRISQUE = ""
      Else
         m_LBRISQUE = theData.fields("LBRISQUE")
      End If

      If IsNull(theData.fields("CDSITUATSIN")) Then
         m_CDSITUATSIN = 0
      Else
         m_CDSITUATSIN = theData.fields("CDSITUATSIN")
      End If

      If IsNull(theData.fields("DTSITUATSIN")) Then
         m_DTSITUATSIN = 0
      Else
         m_DTSITUATSIN = theData.fields("DTSITUATSIN")
      End If

      If IsNull(theData.fields("CDPRETRATTSIN")) Then
         m_CDPRETRATTSIN = ""
      Else
         m_CDPRETRATTSIN = theData.fields("CDPRETRATTSIN")
      End If

      If IsNull(theData.fields("CDCTGPRT")) Then
         m_CDCTGPRT = ""
      Else
         m_CDCTGPRT = theData.fields("CDCTGPRT")
      End If

      If IsNull(theData.fields("LBCTGPRT")) Then
         m_LBCTGPRT = ""
      Else
         m_LBCTGPRT = theData.fields("LBCTGPRT")
      End If

      If IsNull(theData.fields("DTPREECH")) Then
         m_DTPREECH = 0
      Else
         m_DTPREECH = theData.fields("DTPREECH")
      End If

      If IsNull(theData.fields("DTDERECH")) Then
         m_DTDERECH = 0
      Else
         m_DTDERECH = theData.fields("DTDERECH")
      End If

      If IsNull(theData.fields("CDPERIODICITEECH")) Then
         m_CDPERIODICITEECH = ""
      Else
         m_CDPERIODICITEECH = theData.fields("CDPERIODICITEECH")
      End If

      If IsNull(theData.fields("MTECHEANCE1")) Then
         m_MTECHEANCE1 = 0
      Else
         m_MTECHEANCE1 = theData.fields("MTECHEANCE1")
      End If

      If IsNull(theData.fields("DTDEBPERECH1")) Then
         m_DTDEBPERECH1 = 0
      Else
         m_DTDEBPERECH1 = theData.fields("DTDEBPERECH1")
      End If

      If IsNull(theData.fields("DTFINPERECH1")) Then
         m_DTFINPERECH1 = 0
      Else
         m_DTFINPERECH1 = theData.fields("DTFINPERECH1")
      End If

      If IsNull(theData.fields("MTECHEANCE2")) Then
         m_MTECHEANCE2 = 0
      Else
         m_MTECHEANCE2 = theData.fields("MTECHEANCE2")
      End If

      If IsNull(theData.fields("DTDEBPERECH2")) Then
         m_DTDEBPERECH2 = 0
      Else
         m_DTDEBPERECH2 = theData.fields("DTDEBPERECH2")
      End If

      If IsNull(theData.fields("DTFINPERECH2")) Then
         m_DTFINPERECH2 = 0
      Else
         m_DTFINPERECH2 = theData.fields("DTFINPERECH2")
      End If

      If IsNull(theData.fields("MTECHEANCE3")) Then
         m_MTECHEANCE3 = 0
      Else
         m_MTECHEANCE3 = theData.fields("MTECHEANCE3")
      End If

      If IsNull(theData.fields("DTDEBPERECH3")) Then
         m_DTDEBPERECH3 = 0
      Else
         m_DTDEBPERECH3 = theData.fields("DTDEBPERECH3")
      End If

      If IsNull(theData.fields("DTFINPERECH3")) Then
         m_DTFINPERECH3 = 0
      Else
         m_DTFINPERECH3 = theData.fields("DTFINPERECH3")
      End If

      If IsNull(theData.fields("CDTYPAMO")) Then
         m_CDTYPAMO = ""
      Else
         m_CDTYPAMO = theData.fields("CDTYPAMO")
      End If

      If IsNull(theData.fields("LBTYPAMO")) Then
         m_LBTYPAMO = ""
      Else
         m_LBTYPAMO = theData.fields("LBTYPAMO")
      End If

      If IsNull(theData.fields("TXINVPEC")) Then
         m_TXINVPEC = 0
      Else
         m_TXINVPEC = theData.fields("TXINVPEC")
      End If

      If IsNull(theData.fields("DTDEBPERPIP")) Then
         m_DTDEBPERPIP = 0
      Else
         m_DTDEBPERPIP = theData.fields("DTDEBPERPIP")
      End If

      If IsNull(theData.fields("DTFINPERPIP")) Then
         m_DTFINPERPIP = 0
      Else
         m_DTFINPERPIP = theData.fields("DTFINPERPIP")
      End If

      If IsNull(theData.fields("DTSAISIEPERJUSTIF")) Then
         m_DTSAISIEPERJUSTIF = 0
      Else
         m_DTSAISIEPERJUSTIF = theData.fields("DTSAISIEPERJUSTIF")
      End If

      If IsNull(theData.fields("DTDEBPERJUSTIF")) Then
         m_DTDEBPERJUSTIF = 0
      Else
         m_DTDEBPERJUSTIF = theData.fields("DTDEBPERJUSTIF")
      End If

      If IsNull(theData.fields("DTFINPERJUSTIF")) Then
         m_DTFINPERJUSTIF = 0
      Else
         m_DTFINPERJUSTIF = theData.fields("DTFINPERJUSTIF")
      End If

      If IsNull(theData.fields("DTDEBDERPERRGLTADA")) Then
         m_DTDEBDERPERRGLTADA = 0
      Else
         m_DTDEBDERPERRGLTADA = theData.fields("DTDEBDERPERRGLTADA")
      End If

      If IsNull(theData.fields("DTFINDERPERRGLTADA")) Then
         m_DTFINDERPERRGLTADA = 0
      Else
         m_DTFINDERPERRGLTADA = theData.fields("DTFINDERPERRGLTADA")
      End If

      If IsNull(theData.fields("DTDERPERRGLTADA")) Then
         m_DTDERPERRGLTADA = 0
      Else
         m_DTDERPERRGLTADA = theData.fields("DTDERPERRGLTADA")
      End If

      If IsNull(theData.fields("MTDERPERRGLTADA")) Then
         m_MTDERPERRGLTADA = 0
      Else
         m_MTDERPERRGLTADA = theData.fields("MTDERPERRGLTADA")
      End If

      If IsNull(theData.fields("DTDEBDERPERRGLTADC")) Then
         m_DTDEBDERPERRGLTADC = 0
      Else
         m_DTDEBDERPERRGLTADC = theData.fields("DTDEBDERPERRGLTADC")
      End If

      If IsNull(theData.fields("DTFINDERPERRGLTADC")) Then
         m_DTFINDERPERRGLTADC = 0
      Else
         m_DTFINDERPERRGLTADC = theData.fields("DTFINDERPERRGLTADC")
      End If

      If IsNull(theData.fields("DTDERPERRGLTADC")) Then
         m_DTDERPERRGLTADC = 0
      Else
         m_DTDERPERRGLTADC = theData.fields("DTDERPERRGLTADC")
      End If

      If IsNull(theData.fields("MTDERPERRGLTADC")) Then
         m_MTDERPERRGLTADC = 0
      Else
         m_MTDERPERRGLTADC = theData.fields("MTDERPERRGLTADC")
      End If

      If IsNull(theData.fields("CDSINPREPROV")) Then
         m_CDSINPREPROV = ""
      Else
         m_CDSINPREPROV = theData.fields("CDSINPREPROV")
      End If

      If IsNull(theData.fields("MTTOTREGLEICIV")) Then
         m_MTTOTREGLEICIV = 0
      Else
         m_MTTOTREGLEICIV = theData.fields("MTTOTREGLEICIV")
      End If

      If IsNull(theData.fields("DTDEBPROV")) Then
         m_DTDEBPROV = 0
      Else
         m_DTDEBPROV = theData.fields("DTDEBPROV")
      End If

      If IsNull(theData.fields("DTFINPROV")) Then
         m_DTFINPROV = 0
      Else
         m_DTFINPROV = theData.fields("DTFINPROV")
      End If

      If IsNull(theData.fields("INDBASREV")) Then
         m_INDBASREV = ""
      Else
         m_INDBASREV = theData.fields("INDBASREV")
      End If

      If IsNull(theData.fields("MTPREANN")) Then
         m_MTPREANN = 0
      Else
         m_MTPREANN = theData.fields("MTPREANN")
      End If

      If IsNull(theData.fields("MTPREANN2")) Then
         m_MTPREANN2 = 0
      Else
         m_MTPREANN2 = theData.fields("MTPREANN2")
      End If
     
      If bFromP3IPROVCOL = False Then
        If IsNull(theData.fields("MTCPLEST")) Then
           m_MTCPLEST = 0
        Else
           m_MTCPLEST = theData.fields("MTCPLEST")
        End If
  
        If IsNull(theData.fields("MTCPLREE")) Then
           m_MTCPLREE = 0
        Else
           m_MTCPLREE = theData.fields("MTCPLREE")
        End If
  
        If IsNull(theData.fields("MTPROCAL")) Then
           m_MTPROCAL = 0
        Else
           m_MTPROCAL = theData.fields("MTPROCAL")
        End If
  
        If IsNull(theData.fields("MTPROIMP")) Then
           m_MTPROIMP = 0
        Else
           m_MTPROIMP = theData.fields("MTPROIMP")
        End If
  
        If IsNull(theData.fields("MTPROVIT")) Then
           m_MTPROVIT = 0
        Else
           m_MTPROVIT = theData.fields("MTPROVIT")
        End If
  
        If IsNull(theData.fields("MTPROPAS")) Then
           m_MTPROPAS = 0
        Else
           m_MTPROPAS = theData.fields("MTPROPAS")
        End If
  
        If IsNull(theData.fields("TXPROV")) Then
           m_TXPROV = 0
        Else
           m_TXPROV = theData.fields("TXPROV")
        End If
  
        If IsNull(theData.fields("TXPROPASS")) Then
           m_TXPROPASS = 0
        Else
           m_TXPROPASS = theData.fields("TXPROPASS")
        End If
  
        If IsNull(theData.fields("TXTECHN")) Then
           m_TXTECHN = 0
        Else
           m_TXTECHN = theData.fields("TXTECHN")
        End If
  
        If IsNull(theData.fields("TXFRAIS")) Then
           m_TXFRAIS = 0
        Else
           m_TXFRAIS = theData.fields("TXFRAIS")
        End If
  
        If IsNull(theData.fields("MTCAPCONORI")) Then
           m_MTCAPCONORI = 0
        Else
           m_MTCAPCONORI = theData.fields("MTCAPCONORI")
        End If
  
        If IsNull(theData.fields("MTPANCAL")) Then
           m_MTPANCAL = 0
        Else
           m_MTPANCAL = theData.fields("MTPANCAL")
        End If
  
        If IsNull(theData.fields("MTCMPPAN")) Then
           m_MTCMPPAN = 0
        Else
           m_MTCMPPAN = theData.fields("MTCMPPAN")
        End If
  
        If IsNull(theData.fields("MTCOMPAN")) Then
           m_MTCOMPAN = 0
        Else
           m_MTCOMPAN = theData.fields("MTCOMPAN")
        End If
  
        If IsNull(theData.fields("MTCOMCMP")) Then
           m_MTCOMCMP = 0
        Else
           m_MTCOMCMP = theData.fields("MTCOMCMP")
        End If
  
        If IsNull(theData.fields("CDLOTIMPORT")) Then
           m_CDLOTIMPORT = 0
        Else
           m_CDLOTIMPORT = theData.fields("CDLOTIMPORT")
        End If
  
        If IsNull(theData.fields("TYPEPROVISION")) Then
           m_TYPEPROVISION = ""
        Else
           m_TYPEPROVISION = theData.fields("TYPEPROVISION")
        End If
  
        If IsNull(theData.fields("PERALIM")) Then
           m_PERALIM = 0
        Else
           m_PERALIM = theData.fields("PERALIM")
        End If
  
        If IsNull(theData.fields("TXLISSAGE")) Then
           m_TXLISSAGE = 0
        Else
           m_TXLISSAGE = theData.fields("TXLISSAGE")
        End If
  
        If IsNull(theData.fields("TXPROSANSLISSAGE")) Then
           m_TXPROSANSLISSAGE = 0
        Else
           m_TXPROSANSLISSAGE = theData.fields("TXPROSANSLISSAGE")
        End If
  
        ' Evol 2010 - Réforme des retraites
        If IsNull(theData.fields("FLAMORTISSABLE")) Then
           m_FLAMORTISSABLE = 0
        Else
           m_FLAMORTISSABLE = theData.fields("FLAMORTISSABLE")
        End If
        
        If IsNull(theData.fields("TXAMORT")) Then
           m_TXAMORT = 0
        Else
           m_TXAMORT = theData.fields("TXAMORT")
        End If
        
        If IsNull(theData.fields("DTLIMPROAPR")) Then
           m_DTLIMPROAPR = 0
        Else
           m_DTLIMPROAPR = theData.fields("DTLIMPROAPR")
        End If
        
        If IsNull(theData.fields("MTPROIMPAVR")) Then
           m_MTPROIMPAVR = 0
        Else
           m_MTPROIMPAVR = theData.fields("MTPROIMPAVR")
        End If
        
        If IsNull(theData.fields("MTPROIMPAPR")) Then
           m_MTPROIMPAPR = 0
        Else
           m_MTPROIMPAPR = theData.fields("MTPROIMPAPR")
        End If
        
        If IsNull(theData.fields("MTINDEMRES")) Then
           m_MTINDEMRES = 0
        Else
           m_MTINDEMRES = theData.fields("MTINDEMRES")
        End If
        
        If IsNull(theData.fields("MTPROCALAVR")) Then
           m_MTPROCALAVR = 0
        Else
           m_MTPROCALAVR = theData.fields("MTPROCALAVR")
        End If
        
        If IsNull(theData.fields("MTPROCALAPR")) Then
           m_MTPROCALAPR = 0
        Else
           m_MTPROCALAPR = theData.fields("MTPROCALAPR")
        End If
        
        If IsNull(theData.fields("MTCOUTREFPROCAL")) Then
           m_MTCOUTREFPROCAL = 0
        Else
           m_MTCOUTREFPROCAL = theData.fields("MTCOUTREFPROCAL")
        End If
        
        If IsNull(theData.fields("MTAMORTCAL")) Then
           m_MTAMORTCAL = 0
        Else
           m_MTAMORTCAL = theData.fields("MTAMORTCAL")
        End If
        
        If IsNull(theData.fields("MTRESTEAMORCAL")) Then
           m_MTRESTEAMORCAL = 0
        Else
           m_MTRESTEAMORCAL = theData.fields("MTRESTEAMORCAL")
        End If
        
        If IsNull(theData.fields("MTPROPASAVR")) Then
           m_MTPROPASAVR = 0
        Else
           m_MTPROPASAVR = theData.fields("MTPROPASAVR")
        End If
        
        If IsNull(theData.fields("MTPROPASAPR")) Then
           m_MTPROPASAPR = 0
        Else
           m_MTPROPASAPR = theData.fields("MTPROPASAPR")
        End If
        
        If IsNull(theData.fields("MTCOUTREFPROPAS")) Then
           m_MTCOUTREFPROPAS = 0
        Else
           m_MTCOUTREFPROPAS = theData.fields("MTCOUTREFPROPAS")
        End If
        
        If IsNull(theData.fields("MTAMORTPAS")) Then
           m_MTAMORTPAS = 0
        Else
           m_MTAMORTPAS = theData.fields("MTAMORTPAS")
        End If
   
       If IsNull(theData.fields("AGELIMINC")) Then ''
         m_AGELIMINC = ""
       Else
         m_AGELIMINC = theData.fields("AGELIMINC")
       End If
      
       If IsNull(theData.fields("AGELIMINV")) Then ''
         m_AGELIMINV = ""
       Else
         m_AGELIMINV = theData.fields("AGELIMINV")
       End If
    End If
   
    If IsNull(theData.fields("MTPREREV")) Then ''
       m_MTPREREV = 0
    Else
       m_MTPREREV = theData.fields("MTPREREV")
    End If
    
    If IsNull(theData.fields("MTPREMAJ")) Then ''
       m_MTPREMAJ = 0
    Else
       m_MTPREMAJ = theData.fields("MTPREMAJ")
    End If
    
    If IsNull(theData.fields("MTPRIREG")) Then ''
       m_MTPRIREG = 0
    Else
       m_MTPRIREG = theData.fields("MTPRIREG")
    End If
    
    If IsNull(theData.fields("MTPRIRE1")) Then ''
       m_MTPRIRE1 = 0
    Else
       m_MTPRIRE1 = theData.fields("MTPRIRE1")
    End If
    
    If IsNull(theData.fields("MTPRIRE2")) Then ''
       m_MTPRIRE2 = 0
    Else
       m_MTPRIRE2 = theData.fields("MTPRIRE2")
    End If
    
    If IsNull(theData.fields("CDMONNAIE")) Then ''
       m_CDMONNAIE = ""
    Else
       m_CDMONNAIE = theData.fields("CDMONNAIE")
    End If
    
    If IsNull(theData.fields("CDPAYS")) Then ''
       m_CDPAYS = ""
    Else
       m_CDPAYS = theData.fields("CDPAYS")
    End If
    
    If IsNull(theData.fields("CDAPPLISOURCE")) Then ''
       m_CDAPPLISOURCE = 0
    Else
       m_CDAPPLISOURCE = theData.fields("CDAPPLISOURCE")
    End If
    
    If IsNull(theData.fields("LBCOMLIG")) Then ''
       m_LBCOMLIG = ""
    Else
       m_LBCOMLIG = theData.fields("LBCOMLIG")
    End If
    
    If IsNull(theData.fields("NUTRAITP3I")) Then ''
       m_NUTRAITP3I = 0
    Else
       m_NUTRAITP3I = theData.fields("NUTRAITP3I")
    End If
    
    If IsNull(theData.fields("NUENRP3I")) Then ''
       m_NUENRP3I = 0
    Else
       m_NUENRP3I = theData.fields("NUENRP3I")
    End If
    
    
    ' Evol 2010 - Lot 2
    If IsNull(theData.fields("CDCATINV")) Then ''
       m_CDCATINV = ""
    Else
       m_CDCATINV = theData.fields("CDCATINV")
    End If
    
    If IsNull(theData.fields("LBCATINV")) Then ''
       m_LBCATINV = ""
    Else
       m_LBCATINV = theData.fields("LBCATINV")
    End If
    
    If IsNull(theData.fields("CDCONTENTIEUX")) Then ''
       m_CDCONTENTIEUX = ""
    Else
       m_CDCONTENTIEUX = theData.fields("CDCONTENTIEUX")
    End If
    
    If IsNull(theData.fields("NUSINISTRE")) Then ''
       m_NUSINISTRE = ""
    Else
       m_NUSINISTRE = theData.fields("NUSINISTRE")
    End If
    
    If IsNull(theData.fields("CDCHOIXPREST")) Then ''
       m_CDCHOIXPREST = ""
    Else
       m_CDCHOIXPREST = theData.fields("CDCHOIXPREST")
    End If
    
    If IsNull(theData.fields("LBCHOIXPREST")) Then ''
       m_LBCHOIXPREST = ""
    Else
       m_LBCHOIXPREST = theData.fields("LBCHOIXPREST")
    End If
    
    If IsNull(theData.fields("MTCAPCON")) Then
       m_MTCAPCON = 0
    Else
       m_MTCAPCON = theData.fields("MTCAPCON")
    End If
    
    If IsNull(theData.fields("MTCAPSSRISQ")) Then
       m_MTCAPSSRISQ = 0
    Else
       m_MTCAPSSRISQ = theData.fields("MTCAPSSRISQ")
    End If
   
    If IsNull(theData.fields("MTCAPSSRISQ")) Then ''
      m_MTCAPSSRISQ = ""
    Else
      m_MTCAPSSRISQ = theData.fields("MTCAPSSRISQ")
    End If
    
  Else
     ret = False
  End If


  Load_OLD = ret

  Exit Function

err_Load:
  If Err = 3265 Then Resume Next
  MsgBox "Erreur dans clsTTPROVCOLL::Load() : " & Err & vbLf & Err.Description, vbCritical
  Load_OLD = False
  Exit Function
  Resume Next
End Function


'##ModelId=5C8A7600029F
Public Function Load(DB As DataAccess, theData As ADODB.Recordset, bFromP3IPROVCOL As Boolean) As Boolean

   'Dim theData As ADODB.Recordset
   Dim ret As Boolean
   ret = True

   On Error GoTo err_Load
   
   'fieldValue = Replace$(fieldValue, ",", ".")

   If theData.EOF = False Then

      If IsNull(theData.fields("COMMENTAIRE")) Then
         m_COMMENTAIRE = ""
      Else
         m_COMMENTAIRE = Replace$(theData.fields("COMMENTAIRE"), ";", ",")
      End If
      
      If IsNull(theData.fields("CDCOMPAGNIE")) Then
         m_CDCOMPAGNIE = 0
      Else
         m_CDCOMPAGNIE = Replace$(theData.fields("CDCOMPAGNIE"), ",", ".")
      End If

      If IsNull(theData.fields("CDAPPLI")) Then
         m_CDAPPLI = 0
      Else
         m_CDAPPLI = Replace$(theData.fields("CDAPPLI"), ",", ".")
      End If

      If IsNull(theData.fields("CDPRODUIT")) Then
         m_CDPRODUIT = ""
      Else
         m_CDPRODUIT = Replace$(theData.fields("CDPRODUIT"), ",", ".")
      End If

      If IsNull(theData.fields("NUCONTRA")) Then
         m_NUCONTRA = ""
      Else
         m_NUCONTRA = Replace$(theData.fields("NUCONTRA"), ",", ".")
      End If

      If IsNull(theData.fields("NUPRET")) Then
         m_NUPRET = 0
      Else
         m_NUPRET = Replace$(theData.fields("NUPRET"), ",", ".")
      End If

      If IsNull(theData.fields("NUDOSSIERPREST")) Then
         m_NUDOSSIERPREST = ""
      Else
         m_NUDOSSIERPREST = Replace$(theData.fields("NUDOSSIERPREST"), ",", ".")
      End If

      If IsNull(theData.fields("NUSOUSDOSSIERPREST")) Then
         m_NUSOUSDOSSIERPREST = ""
      Else
         m_NUSOUSDOSSIERPREST = Replace$(theData.fields("NUSOUSDOSSIERPREST"), ",", ".")
      End If

      If IsNull(theData.fields("NUCONTRATGESTDELEG")) Then
         m_NUCONTRATGESTDELEG = ""
      Else
         m_NUCONTRATGESTDELEG = Replace$(theData.fields("NUCONTRATGESTDELEG"), ",", ".")
      End If

      If IsNull(theData.fields("CDGARAN")) Then
         m_CDGARAN = 0
      Else
         m_CDGARAN = Replace$(theData.fields("CDGARAN"), ",", ".")
      End If

      If IsNull(theData.fields("NUREFMVT")) Then
         m_NUREFMVT = 0
      Else
         m_NUREFMVT = Replace$(theData.fields("NUREFMVT"), ",", ".")
      End If

      If IsNull(theData.fields("CDTYPMVT")) Then
         m_CDTYPMVT = 0
      Else
         m_CDTYPMVT = Replace$(theData.fields("CDTYPMVT"), ",", ".")
      End If

      If IsNull(theData.fields("NUTRAITE")) Then
         m_NUTRAITE = 0
      Else
         m_NUTRAITE = Replace$(theData.fields("NUTRAITE"), ",", ".")
      End If

      If IsNull(theData.fields("CDCATADHESION")) Then
         m_CDCATADHESION = 0
      Else
         m_CDCATADHESION = Replace$(theData.fields("CDCATADHESION"), ",", ".")
      End If

      If IsNull(theData.fields("LBCATADHESION")) Then
         m_LBCATADHESION = ""
      Else
         m_LBCATADHESION = Replace$(theData.fields("LBCATADHESION"), ",", ".")
      End If

      If IsNull(theData.fields("CDOPTION")) Then
         m_CDOPTION = ""
      Else
         m_CDOPTION = Replace$(theData.fields("CDOPTION"), ",", ".")
      End If

      If IsNull(theData.fields("LBOPTIONCON")) Then
         m_LBOPTIONCON = ""
      Else
         m_LBOPTIONCON = Replace$(theData.fields("LBOPTIONCON"), ",", ".")
      End If

      If IsNull(theData.fields("LBSOUSCR")) Then
         m_LBSOUSCR = ""
      Else
         m_LBSOUSCR = Replace$(theData.fields("LBSOUSCR"), ",", ".")
      End If

      If IsNull(theData.fields("IDASSUREAGI")) Then
         m_IDASSUREAGI = ""
      Else
         m_IDASSUREAGI = Replace$(theData.fields("IDASSUREAGI"), ",", ".")
      End If

      If IsNull(theData.fields("IDASSURE")) Then
         m_IDASSURE = 0
      Else
         m_IDASSURE = Replace$(theData.fields("IDASSURE"), ",", ".")
      End If

      If IsNull(theData.fields("LBASSURE")) Then
         m_LBASSURE = ""
      Else
         m_LBASSURE = Replace$(theData.fields("LBASSURE"), ",", ".")
      End If

      If IsNull(theData.fields("DTNAISSASS")) Then
         m_DTNAISSASS = 0
      Else
         m_DTNAISSASS = Replace$(theData.fields("DTNAISSASS"), ",", ".")
      End If

      If IsNull(theData.fields("CDSEXASSURE")) Then
         m_CDSEXASSURE = ""
      Else
         m_CDSEXASSURE = Replace$(theData.fields("CDSEXASSURE"), ",", ".")
      End If

      If IsNull(theData.fields("IDRENTIERAGI")) Then
         m_IDRENTIERAGI = ""
      Else
         m_IDRENTIERAGI = Replace$(theData.fields("IDRENTIERAGI"), ",", ".")
      End If

      If IsNull(theData.fields("IDRENTIER")) Then
         m_IDRENTIER = 0
      Else
         m_IDRENTIER = Replace$(theData.fields("IDRENTIER"), ",", ".")
      End If

      If IsNull(theData.fields("LBRENTIER")) Then
         m_LBRENTIER = ""
      Else
         m_LBRENTIER = Replace$(theData.fields("LBRENTIER"), ",", ".")
      End If

      If IsNull(theData.fields("DTNAISSREN")) Then
         m_DTNAISSREN = 0
      Else
         m_DTNAISSREN = Replace$(theData.fields("DTNAISSREN"), ",", ".")
      End If

      If IsNull(theData.fields("IDCORENTIERAGI")) Then
         m_IDCORENTIERAGI = ""
      Else
         m_IDCORENTIERAGI = Replace$(theData.fields("IDCORENTIERAGI"), ",", ".")
      End If

      If IsNull(theData.fields("IDCORENTIER")) Then
         m_IDCORENTIER = 0
      Else
         m_IDCORENTIER = Replace$(theData.fields("IDCORENTIER"), ",", ".")
      End If

      If IsNull(theData.fields("LBCORENTIER")) Then
         m_LBCORENTIER = ""
      Else
         m_LBCORENTIER = Replace$(theData.fields("LBCORENTIER"), ",", ".")
      End If

      If IsNull(theData.fields("DTNAISSCOR")) Then
         m_DTNAISSCOR = 0
      Else
         m_DTNAISSCOR = Replace$(theData.fields("DTNAISSCOR"), ",", ".")
      End If

      If IsNull(theData.fields("DTSURVSIN")) Then
         m_DTSURVSIN = 0
      Else
         m_DTSURVSIN = Replace$(theData.fields("DTSURVSIN"), ",", ".")
      End If

      If IsNull(theData.fields("AGESURVSIN")) Then
         m_AGESURVSIN = 0
      Else
         m_AGESURVSIN = Replace$(theData.fields("AGESURVSIN"), ",", ".")
      End If

      If IsNull(theData.fields("DURPOSARR")) Then
         m_DURPOSARR = 0
      Else
         m_DURPOSARR = Replace$(theData.fields("DURPOSARR"), ",", ".")
      End If

      If IsNull(theData.fields("ANCARRTRA")) Then
         m_ANCARRTRA = 0
      Else
         m_ANCARRTRA = Replace$(theData.fields("ANCARRTRA"), ",", ".")
      End If

      If IsNull(theData.fields("CDPERIODICITE")) Then
         m_CDPERIODICITE = ""
      Else
         m_CDPERIODICITE = Replace$(theData.fields("CDPERIODICITE"), ",", ".")
      End If

      If IsNull(theData.fields("CDTYPTERME")) Then
         m_CDTYPTERME = 0
      Else
         m_CDTYPTERME = Replace$(theData.fields("CDTYPTERME"), ",", ".")
      End If

      If IsNull(theData.fields("DTEFFREN")) Then
         m_DTEFFREN = 0
      Else
         m_DTEFFREN = Replace$(theData.fields("DTEFFREN"), ",", ".")
      End If

      If IsNull(theData.fields("DTLIMPROAVR")) Then
         m_DTLIMPROAVR = 0
      Else
         m_DTLIMPROAVR = Replace$(theData.fields("DTLIMPROAVR"), ",", ".")
      End If

      If IsNull(theData.fields("DTDERREG")) Then
         m_DTDERREG = 0
      Else
         m_DTDERREG = Replace$(theData.fields("DTDERREG"), ",", ".")
      End If

      If IsNull(theData.fields("DTDEBPER")) Then
         m_DTDEBPER = 0
      Else
         m_DTDEBPER = Replace$(theData.fields("DTDEBPER"), ",", ".")
      End If

      If IsNull(theData.fields("DTFINPER")) Then
         m_DTFINPER = 0
      Else
         m_DTFINPER = Replace$(theData.fields("DTFINPER"), ",", ".")
      End If

      If IsNull(theData.fields("CDPROEVA")) Then
         m_CDPROEVA = ""
      Else
         m_CDPROEVA = Replace$(theData.fields("CDPROEVA"), ",", ".")
      End If

      If IsNull(theData.fields("ANNEADHE")) Then
         m_ANNEADHE = 0
      Else
         m_ANNEADHE = Replace$(theData.fields("ANNEADHE"), ",", ".")
      End If

      If IsNull(theData.fields("CDSINCON")) Then
         m_CDSINCON = ""
      Else
         m_CDSINCON = Replace$(theData.fields("CDSINCON"), ",", ".")
      End If

      If IsNull(theData.fields("CDMISINV")) Then
         m_CDMISINV = ""
      Else
         m_CDMISINV = Replace$(theData.fields("CDMISINV"), ",", ".")
      End If

      If IsNull(theData.fields("DTMISINV")) Then
         m_DTMISINV = 0
      Else
         m_DTMISINV = Replace$(theData.fields("DTMISINV"), ",", ".")
      End If

      If IsNull(theData.fields("NUETABLI")) Then
         m_NUETABLI = 0
      Else
         m_NUETABLI = Replace$(theData.fields("NUETABLI"), ",", ".")
      End If

      If IsNull(theData.fields("TXREVERSION")) Then
         m_TXREVERSION = 0
      Else
         m_TXREVERSION = Replace$(theData.fields("TXREVERSION"), ",", ".")
      End If

      If IsNull(theData.fields("DTCALCULPROV")) Then
         m_DTCALCULPROV = 0
      Else
         m_DTCALCULPROV = Replace$(theData.fields("DTCALCULPROV"), ",", ".")
      End If

      If IsNull(theData.fields("DTTRAITPROV")) Then
         m_DTTRAITPROV = 0
      Else
         m_DTTRAITPROV = Replace$(theData.fields("DTTRAITPROV"), ",", ".")
      End If

      If IsNull(theData.fields("DTCREATI")) Then
         m_DTCREATI = 0
      Else
         m_DTCREATI = Replace$(theData.fields("DTCREATI"), ",", ".")
      End If

      If IsNull(theData.fields("DTSIGBIA")) Then
         m_DTSIGBIA = 0
      Else
         m_DTSIGBIA = Replace$(theData.fields("DTSIGBIA"), ",", ".")
      End If

      If IsNull(theData.fields("DTDECSIN")) Then
         m_DTDECSIN = 0
      Else
         m_DTDECSIN = Replace$(theData.fields("DTDECSIN"), ",", ".")
      End If

      If IsNull(theData.fields("CDRISQUE")) Then
         m_CDRISQUE = ""
      Else
         m_CDRISQUE = Replace$(theData.fields("CDRISQUE"), ",", ".")
      End If

      If IsNull(theData.fields("LBRISQUE")) Then
         m_LBRISQUE = ""
      Else
         m_LBRISQUE = Replace$(theData.fields("LBRISQUE"), ",", ".")
      End If

      If IsNull(theData.fields("CDSITUATSIN")) Then
         m_CDSITUATSIN = 0
      Else
         m_CDSITUATSIN = Replace$(theData.fields("CDSITUATSIN"), ",", ".")
      End If

      If IsNull(theData.fields("DTSITUATSIN")) Then
         m_DTSITUATSIN = 0
      Else
         m_DTSITUATSIN = Replace$(theData.fields("DTSITUATSIN"), ",", ".")
      End If

      If IsNull(theData.fields("CDPRETRATTSIN")) Then
         m_CDPRETRATTSIN = ""
      Else
         m_CDPRETRATTSIN = Replace$(theData.fields("CDPRETRATTSIN"), ",", ".")
      End If

      If IsNull(theData.fields("CDCTGPRT")) Then
         m_CDCTGPRT = ""
      Else
         m_CDCTGPRT = Replace$(theData.fields("CDCTGPRT"), ",", ".")
      End If

      If IsNull(theData.fields("LBCTGPRT")) Then
         m_LBCTGPRT = ""
      Else
         m_LBCTGPRT = Replace$(theData.fields("LBCTGPRT"), ",", ".")
      End If

      If IsNull(theData.fields("DTPREECH")) Then
         m_DTPREECH = 0
      Else
         m_DTPREECH = Replace$(theData.fields("DTPREECH"), ",", ".")
      End If

      If IsNull(theData.fields("DTDERECH")) Then
         m_DTDERECH = 0
      Else
         m_DTDERECH = Replace$(theData.fields("DTDERECH"), ",", ".")
      End If

      If IsNull(theData.fields("CDPERIODICITEECH")) Then
         m_CDPERIODICITEECH = ""
      Else
         m_CDPERIODICITEECH = Replace$(theData.fields("CDPERIODICITEECH"), ",", ".")
      End If

      If IsNull(theData.fields("MTECHEANCE1")) Then
         m_MTECHEANCE1 = 0
      Else
         m_MTECHEANCE1 = Replace$(theData.fields("MTECHEANCE1"), ",", ".")
      End If

      If IsNull(theData.fields("DTDEBPERECH1")) Then
         m_DTDEBPERECH1 = 0
      Else
         m_DTDEBPERECH1 = Replace$(theData.fields("DTDEBPERECH1"), ",", ".")
      End If

      If IsNull(theData.fields("DTFINPERECH1")) Then
         m_DTFINPERECH1 = 0
      Else
         m_DTFINPERECH1 = Replace$(theData.fields("DTFINPERECH1"), ",", ".")
      End If

      If IsNull(theData.fields("MTECHEANCE2")) Then
         m_MTECHEANCE2 = 0
      Else
         m_MTECHEANCE2 = Replace$(theData.fields("MTECHEANCE2"), ",", ".")
      End If

      If IsNull(theData.fields("DTDEBPERECH2")) Then
         m_DTDEBPERECH2 = 0
      Else
         m_DTDEBPERECH2 = Replace$(theData.fields("DTDEBPERECH2"), ",", ".")
      End If

      If IsNull(theData.fields("DTFINPERECH2")) Then
         m_DTFINPERECH2 = 0
      Else
         m_DTFINPERECH2 = Replace$(theData.fields("DTFINPERECH2"), ",", ".")
      End If

      If IsNull(theData.fields("MTECHEANCE3")) Then
         m_MTECHEANCE3 = 0
      Else
         m_MTECHEANCE3 = Replace$(theData.fields("MTECHEANCE3"), ",", ".")
      End If

      If IsNull(theData.fields("DTDEBPERECH3")) Then
         m_DTDEBPERECH3 = 0
      Else
         m_DTDEBPERECH3 = Replace$(theData.fields("DTDEBPERECH3"), ",", ".")
      End If

      If IsNull(theData.fields("DTFINPERECH3")) Then
         m_DTFINPERECH3 = 0
      Else
         m_DTFINPERECH3 = Replace$(theData.fields("DTFINPERECH3"), ",", ".")
      End If

      If IsNull(theData.fields("CDTYPAMO")) Then
         m_CDTYPAMO = ""
      Else
         m_CDTYPAMO = Replace$(theData.fields("CDTYPAMO"), ",", ".")
      End If

      If IsNull(theData.fields("LBTYPAMO")) Then
         m_LBTYPAMO = ""
      Else
         m_LBTYPAMO = Replace$(theData.fields("LBTYPAMO"), ",", ".")
      End If

      If IsNull(theData.fields("TXINVPEC")) Then
         m_TXINVPEC = 0
      Else
         m_TXINVPEC = Replace$(theData.fields("TXINVPEC"), ",", ".")
      End If

      If IsNull(theData.fields("DTDEBPERPIP")) Then
         m_DTDEBPERPIP = 0
      Else
         m_DTDEBPERPIP = Replace$(theData.fields("DTDEBPERPIP"), ",", ".")
      End If

      If IsNull(theData.fields("DTFINPERPIP")) Then
         m_DTFINPERPIP = 0
      Else
         m_DTFINPERPIP = Replace$(theData.fields("DTFINPERPIP"), ",", ".")
      End If

      If IsNull(theData.fields("DTSAISIEPERJUSTIF")) Then
         m_DTSAISIEPERJUSTIF = 0
      Else
         m_DTSAISIEPERJUSTIF = Replace$(theData.fields("DTSAISIEPERJUSTIF"), ",", ".")
      End If

      If IsNull(theData.fields("DTDEBPERJUSTIF")) Then
         m_DTDEBPERJUSTIF = 0
      Else
         m_DTDEBPERJUSTIF = Replace$(theData.fields("DTDEBPERJUSTIF"), ",", ".")
      End If

      If IsNull(theData.fields("DTFINPERJUSTIF")) Then
         m_DTFINPERJUSTIF = 0
      Else
         m_DTFINPERJUSTIF = Replace$(theData.fields("DTFINPERJUSTIF"), ",", ".")
      End If

      If IsNull(theData.fields("DTDEBDERPERRGLTADA")) Then
         m_DTDEBDERPERRGLTADA = 0
      Else
         m_DTDEBDERPERRGLTADA = Replace$(theData.fields("DTDEBDERPERRGLTADA"), ",", ".")
      End If

      If IsNull(theData.fields("DTFINDERPERRGLTADA")) Then
         m_DTFINDERPERRGLTADA = 0
      Else
         m_DTFINDERPERRGLTADA = Replace$(theData.fields("DTFINDERPERRGLTADA"), ",", ".")
      End If

      If IsNull(theData.fields("DTDERPERRGLTADA")) Then
         m_DTDERPERRGLTADA = 0
      Else
         m_DTDERPERRGLTADA = Replace$(theData.fields("DTDERPERRGLTADA"), ",", ".")
      End If

      If IsNull(theData.fields("MTDERPERRGLTADA")) Then
         m_MTDERPERRGLTADA = 0
      Else
         m_MTDERPERRGLTADA = Replace$(theData.fields("MTDERPERRGLTADA"), ",", ".")
      End If

      If IsNull(theData.fields("DTDEBDERPERRGLTADC")) Then
         m_DTDEBDERPERRGLTADC = 0
      Else
         m_DTDEBDERPERRGLTADC = Replace$(theData.fields("DTDEBDERPERRGLTADC"), ",", ".")
      End If

      If IsNull(theData.fields("DTFINDERPERRGLTADC")) Then
         m_DTFINDERPERRGLTADC = 0
      Else
         m_DTFINDERPERRGLTADC = Replace$(theData.fields("DTFINDERPERRGLTADC"), ",", ".")
      End If

      If IsNull(theData.fields("DTDERPERRGLTADC")) Then
         m_DTDERPERRGLTADC = 0
      Else
         m_DTDERPERRGLTADC = Replace$(theData.fields("DTDERPERRGLTADC"), ",", ".")
      End If

      If IsNull(theData.fields("MTDERPERRGLTADC")) Then
         m_MTDERPERRGLTADC = 0
      Else
         m_MTDERPERRGLTADC = Replace$(theData.fields("MTDERPERRGLTADC"), ",", ".")
      End If

      If IsNull(theData.fields("CDSINPREPROV")) Then
         m_CDSINPREPROV = ""
      Else
         m_CDSINPREPROV = Replace$(theData.fields("CDSINPREPROV"), ",", ".")
      End If

      If IsNull(theData.fields("MTTOTREGLEICIV")) Then
         m_MTTOTREGLEICIV = 0
      Else
         m_MTTOTREGLEICIV = Replace$(theData.fields("MTTOTREGLEICIV"), ",", ".")
      End If

      If IsNull(theData.fields("DTDEBPROV")) Then
         m_DTDEBPROV = 0
      Else
         m_DTDEBPROV = Replace$(theData.fields("DTDEBPROV"), ",", ".")
      End If

      If IsNull(theData.fields("DTFINPROV")) Then
         m_DTFINPROV = 0
      Else
         m_DTFINPROV = Replace$(theData.fields("DTFINPROV"), ",", ".")
      End If

      If IsNull(theData.fields("INDBASREV")) Then
         m_INDBASREV = ""
      Else
         m_INDBASREV = Replace$(theData.fields("INDBASREV"), ",", ".")
      End If

      If IsNull(theData.fields("MTPREANN")) Then
         m_MTPREANN = 0
      Else
         m_MTPREANN = Replace$(theData.fields("MTPREANN"), ",", ".")
      End If

      If IsNull(theData.fields("MTPREANN2")) Then
         m_MTPREANN2 = 0
      Else
         m_MTPREANN2 = Replace$(theData.fields("MTPREANN2"), ",", ".")
      End If
     
      If bFromP3IPROVCOL = False Then
        If IsNull(theData.fields("MTCPLEST")) Then
           m_MTCPLEST = 0
        Else
           m_MTCPLEST = Replace$(theData.fields("MTCPLEST"), ",", ".")
        End If
  
        If IsNull(theData.fields("MTCPLREE")) Then
           m_MTCPLREE = 0
        Else
           m_MTCPLREE = Replace$(theData.fields("MTCPLREE"), ",", ".")
        End If
  
        If IsNull(theData.fields("MTPROCAL")) Then
           m_MTPROCAL = 0
        Else
           m_MTPROCAL = Replace$(theData.fields("MTPROCAL"), ",", ".")
        End If
  
        If IsNull(theData.fields("MTPROIMP")) Then
           m_MTPROIMP = 0
        Else
           m_MTPROIMP = Replace$(theData.fields("MTPROIMP"), ",", ".")
        End If
  
        If IsNull(theData.fields("MTPROVIT")) Then
           m_MTPROVIT = 0
        Else
           m_MTPROVIT = Replace$(theData.fields("MTPROVIT"), ",", ".")
        End If
  
        If IsNull(theData.fields("MTPROPAS")) Then
           m_MTPROPAS = 0
        Else
           m_MTPROPAS = Replace$(theData.fields("MTPROPAS"), ",", ".")
        End If
  
        If IsNull(theData.fields("TXPROV")) Then
           m_TXPROV = 0
        Else
           m_TXPROV = Replace$(theData.fields("TXPROV"), ",", ".")
        End If
  
        If IsNull(theData.fields("TXPROPASS")) Then
           m_TXPROPASS = 0
        Else
           m_TXPROPASS = Replace$(theData.fields("TXPROPASS"), ",", ".")
        End If
  
        If IsNull(theData.fields("TXTECHN")) Then
           m_TXTECHN = 0
        Else
           m_TXTECHN = Replace$(theData.fields("TXTECHN"), ",", ".")
        End If
  
        If IsNull(theData.fields("TXFRAIS")) Then
           m_TXFRAIS = 0
        Else
           m_TXFRAIS = Replace$(theData.fields("TXFRAIS"), ",", ".")
        End If
  
        If IsNull(theData.fields("MTCAPCONORI")) Then
           m_MTCAPCONORI = 0
        Else
           m_MTCAPCONORI = Replace$(theData.fields("MTCAPCONORI"), ",", ".")
        End If
  
        If IsNull(theData.fields("MTPANCAL")) Then
           m_MTPANCAL = 0
        Else
           m_MTPANCAL = Replace$(theData.fields("MTPANCAL"), ",", ".")
        End If
  
        If IsNull(theData.fields("MTCMPPAN")) Then
           m_MTCMPPAN = 0
        Else
           m_MTCMPPAN = Replace$(theData.fields("MTCMPPAN"), ",", ".")
        End If
  
        If IsNull(theData.fields("MTCOMPAN")) Then
           m_MTCOMPAN = 0
        Else
           m_MTCOMPAN = Replace$(theData.fields("MTCOMPAN"), ",", ".")
        End If
  
        If IsNull(theData.fields("MTCOMCMP")) Then
           m_MTCOMCMP = 0
        Else
           m_MTCOMCMP = Replace$(theData.fields("MTCOMCMP"), ",", ".")
        End If
  
        If IsNull(theData.fields("CDLOTIMPORT")) Then
           m_CDLOTIMPORT = 0
        Else
           m_CDLOTIMPORT = Replace$(theData.fields("CDLOTIMPORT"), ",", ".")
        End If
  
        If IsNull(theData.fields("TYPEPROVISION")) Then
           m_TYPEPROVISION = ""
        Else
           m_TYPEPROVISION = Replace$(theData.fields("TYPEPROVISION"), ",", ".")
        End If
  
        If IsNull(theData.fields("PERALIM")) Then
           m_PERALIM = 0
        Else
           m_PERALIM = Replace$(theData.fields("PERALIM"), ",", ".")
        End If
  
        If IsNull(theData.fields("TXLISSAGE")) Then
           m_TXLISSAGE = 0
        Else
           m_TXLISSAGE = Replace$(theData.fields("TXLISSAGE"), ",", ".")
        End If
  
        If IsNull(theData.fields("TXPROSANSLISSAGE")) Then
           m_TXPROSANSLISSAGE = 0
        Else
           m_TXPROSANSLISSAGE = Replace$(theData.fields("TXPROSANSLISSAGE"), ",", ".")
        End If
  
        ' Evol 2010 - Réforme des retraites
        If IsNull(theData.fields("FLAMORTISSABLE")) Then
           m_FLAMORTISSABLE = 0
        Else
           m_FLAMORTISSABLE = Replace$(theData.fields("FLAMORTISSABLE"), ",", ".")
        End If
        
        If IsNull(theData.fields("TXAMORT")) Then
           m_TXAMORT = 0
        Else
           m_TXAMORT = Replace$(theData.fields("TXAMORT"), ",", ".")
        End If
        
        If IsNull(theData.fields("DTLIMPROAPR")) Then
           m_DTLIMPROAPR = 0
        Else
           m_DTLIMPROAPR = Replace$(theData.fields("DTLIMPROAPR"), ",", ".")
        End If
        
        If IsNull(theData.fields("MTPROIMPAVR")) Then
           m_MTPROIMPAVR = 0
        Else
           m_MTPROIMPAVR = Replace$(theData.fields("MTPROIMPAVR"), ",", ".")
        End If
        
        If IsNull(theData.fields("MTPROIMPAPR")) Then
           m_MTPROIMPAPR = 0
        Else
           m_MTPROIMPAPR = Replace$(theData.fields("MTPROIMPAPR"), ",", ".")
        End If
        
        If IsNull(theData.fields("MTINDEMRES")) Then
           m_MTINDEMRES = 0
        Else
           m_MTINDEMRES = Replace$(theData.fields("MTINDEMRES"), ",", ".")
        End If
        
        If IsNull(theData.fields("MTPROCALAVR")) Then
           m_MTPROCALAVR = 0
        Else
           m_MTPROCALAVR = Replace$(theData.fields("MTPROCALAVR"), ",", ".")
        End If
        
        If IsNull(theData.fields("MTPROCALAPR")) Then
           m_MTPROCALAPR = 0
        Else
           m_MTPROCALAPR = Replace$(theData.fields("MTPROCALAPR"), ",", ".")
        End If
        
        If IsNull(theData.fields("MTCOUTREFPROCAL")) Then
           m_MTCOUTREFPROCAL = 0
        Else
           m_MTCOUTREFPROCAL = Replace$(theData.fields("MTCOUTREFPROCAL"), ",", ".")
        End If
        
        If IsNull(theData.fields("MTAMORTCAL")) Then
           m_MTAMORTCAL = 0
        Else
           m_MTAMORTCAL = Replace$(theData.fields("MTAMORTCAL"), ",", ".")
        End If
        
        If IsNull(theData.fields("MTRESTEAMORCAL")) Then
           m_MTRESTEAMORCAL = 0
        Else
           m_MTRESTEAMORCAL = Replace$(theData.fields("MTRESTEAMORCAL"), ",", ".")
        End If
        
        If IsNull(theData.fields("MTPROPASAVR")) Then
           m_MTPROPASAVR = 0
        Else
           m_MTPROPASAVR = Replace$(theData.fields("MTPROPASAVR"), ",", ".")
        End If
        
        If IsNull(theData.fields("MTPROPASAPR")) Then
           m_MTPROPASAPR = 0
        Else
           m_MTPROPASAPR = Replace$(theData.fields("MTPROPASAPR"), ",", ".")
        End If
        
        If IsNull(theData.fields("MTCOUTREFPROPAS")) Then
           m_MTCOUTREFPROPAS = 0
        Else
           m_MTCOUTREFPROPAS = Replace$(theData.fields("MTCOUTREFPROPAS"), ",", ".")
        End If
        
        If IsNull(theData.fields("MTAMORTPAS")) Then
           m_MTAMORTPAS = 0
        Else
           m_MTAMORTPAS = Replace$(theData.fields("MTAMORTPAS"), ",", ".")
        End If
   
       If IsNull(theData.fields("AGELIMINC")) Then ''
         m_AGELIMINC = ""
       Else
         m_AGELIMINC = Replace$(theData.fields("AGELIMINC"), ",", ".")
       End If
      
       If IsNull(theData.fields("AGELIMINV")) Then ''
         m_AGELIMINV = ""
       Else
         m_AGELIMINV = Replace$(theData.fields("AGELIMINV"), ",", ".")
       End If
    End If
   
    If IsNull(theData.fields("MTPREREV")) Then ''
       m_MTPREREV = 0
    Else
       m_MTPREREV = Replace$(theData.fields("MTPREREV"), ",", ".")
    End If
    
    If IsNull(theData.fields("MTPREMAJ")) Then ''
       m_MTPREMAJ = 0
    Else
       m_MTPREMAJ = Replace$(theData.fields("MTPREMAJ"), ",", ".")
    End If
    
    If IsNull(theData.fields("MTPRIREG")) Then ''
       m_MTPRIREG = 0
    Else
       m_MTPRIREG = Replace$(theData.fields("MTPRIREG"), ",", ".")
    End If
    
    If IsNull(theData.fields("MTPRIRE1")) Then ''
       m_MTPRIRE1 = 0
    Else
       m_MTPRIRE1 = Replace$(theData.fields("MTPRIRE1"), ",", ".")
    End If
    
    If IsNull(theData.fields("MTPRIRE2")) Then ''
       m_MTPRIRE2 = 0
    Else
       m_MTPRIRE2 = Replace$(theData.fields("MTPRIRE2"), ",", ".")
    End If
    
    If IsNull(theData.fields("CDMONNAIE")) Then ''
       m_CDMONNAIE = ""
    Else
       m_CDMONNAIE = Replace$(theData.fields("CDMONNAIE"), ",", ".")
    End If
    
    If IsNull(theData.fields("CDPAYS")) Then ''
       m_CDPAYS = ""
    Else
       m_CDPAYS = Replace$(theData.fields("CDPAYS"), ",", ".")
    End If
    
    If IsNull(theData.fields("CDAPPLISOURCE")) Then ''
       m_CDAPPLISOURCE = 0
    Else
       m_CDAPPLISOURCE = Replace$(theData.fields("CDAPPLISOURCE"), ",", ".")
    End If
    
    If IsNull(theData.fields("LBCOMLIG")) Then ''
       m_LBCOMLIG = ""
    Else
       m_LBCOMLIG = Replace$(theData.fields("LBCOMLIG"), ",", ".")
    End If
    
    If IsNull(theData.fields("NUTRAITP3I")) Then ''
       m_NUTRAITP3I = 0
    Else
       m_NUTRAITP3I = Replace$(theData.fields("NUTRAITP3I"), ",", ".")
    End If
    
    If IsNull(theData.fields("NUENRP3I")) Then ''
       m_NUENRP3I = 0
    Else
       m_NUENRP3I = Replace$(theData.fields("NUENRP3I"), ",", ".")
    End If
    
    
    ' Evol 2010 - Lot 2
    If IsNull(theData.fields("CDCATINV")) Then ''
       m_CDCATINV = ""
    Else
       m_CDCATINV = Replace$(theData.fields("CDCATINV"), ",", ".")
    End If
    
    If IsNull(theData.fields("LBCATINV")) Then ''
       m_LBCATINV = ""
    Else
       m_LBCATINV = Replace$(theData.fields("LBCATINV"), ",", ".")
    End If
    
    If IsNull(theData.fields("CDCONTENTIEUX")) Then ''
       m_CDCONTENTIEUX = ""
    Else
       m_CDCONTENTIEUX = Replace$(theData.fields("CDCONTENTIEUX"), ",", ".")
    End If
    
    If IsNull(theData.fields("NUSINISTRE")) Then ''
       m_NUSINISTRE = ""
    Else
       m_NUSINISTRE = Replace$(theData.fields("NUSINISTRE"), ",", ".")
    End If
    
    If IsNull(theData.fields("CDCHOIXPREST")) Then ''
       m_CDCHOIXPREST = ""
    Else
       m_CDCHOIXPREST = Replace$(theData.fields("CDCHOIXPREST"), ",", ".")
    End If
    
    If IsNull(theData.fields("LBCHOIXPREST")) Then ''
       m_LBCHOIXPREST = ""
    Else
       m_LBCHOIXPREST = Replace$(theData.fields("LBCHOIXPREST"), ",", ".")
    End If
    
    If IsNull(theData.fields("MTCAPCON")) Then
       m_MTCAPCON = 0
    Else
       m_MTCAPCON = Replace$(theData.fields("MTCAPCON"), ",", ".")
    End If
    
    If IsNull(theData.fields("MTCAPSSRISQ")) Then
       m_MTCAPSSRISQ = 0
    Else
       m_MTCAPSSRISQ = Replace$(theData.fields("MTCAPSSRISQ"), ",", ".")
    End If
   
    If IsNull(theData.fields("MTCAPSSRISQ")) Then ''
      m_MTCAPSSRISQ = ""
    Else
      m_MTCAPSSRISQ = Replace$(theData.fields("MTCAPSSRISQ"), ",", ".")
    End If
    
  Else
     ret = False
  End If


  Load = ret

  Exit Function

err_Load:
  If Err = 3265 Then Resume Next
  MsgBox "Erreur dans clsTTPROVCOLL::Load() : " & Err & vbLf & Err.Description, vbCritical
  Load = False
  Exit Function
  Resume Next
End Function

'##ModelId=5C8A76000317
Public Function SaveToCSV()

Dim line As String
 
  line = m_CDCOMPAGNIE & ";" & m_CDAPPLI & ";" & m_CDPRODUIT & ";" & m_NUCONTRA & ";" & m_NUPRET & ";" & m_NUDOSSIERPREST & ";" & m_NUSOUSDOSSIERPREST & ";" & m_NUCONTRATGESTDELEG & ";" & _
m_CDGARAN & ";" & m_NUREFMVT & ";" & m_CDTYPMVT & ";" & m_NUTRAITE & ";" & m_CDCATADHESION & ";" & m_LBCATADHESION & ";" & m_CDOPTION & ";" & m_LBOPTIONCON & ";" & m_LBSOUSCR & ";"

line = line & m_IDASSUREAGI & ";" & m_IDASSURE & ";" & m_LBASSURE & ";" & m_DTNAISSASS & ";" & m_CDSEXASSURE & ";" & m_IDRENTIERAGI & ";" & _
m_IDRENTIER & ";" & m_LBRENTIER & ";" & m_DTNAISSREN & ";" & m_IDCORENTIERAGI & ";" & m_IDCORENTIER & ";" & m_LBCORENTIER & ";" & m_DTNAISSCOR & ";" & m_DTSURVSIN & ";" & m_AGESURVSIN & ";" _

line = line & m_DURPOSARR & ";" & m_ANCARRTRA & ";" & _
m_CDPERIODICITE & ";" & m_CDTYPTERME & ";" & m_DTEFFREN & ";" & m_DTLIMPROAVR & ";" & m_DTDERREG & ";" & m_DTDEBPER & ";" & m_DTFINPER & ";" & m_CDPROEVA & ";" & m_ANNEADHE & ";" & _
m_CDSINCON & ";" & m_CDMISINV & ";" & m_DTMISINV & ";" & m_NUETABLI & ";" & m_TXREVERSION & ";" & m_DTCALCULPROV & ";" & m_DTTRAITPROV & ";" & m_DTCREATI & ";" & m_DTSIGBIA & ";" & _
m_DTDECSIN & ";" & m_CDRISQUE & ";" & m_LBRISQUE & ";" & m_CDSITUATSIN & ";" & m_DTSITUATSIN & ";" & m_CDPRETRATTSIN & ";" & m_CDCTGPRT & ";" & m_LBCTGPRT & ";" & m_DTPREECH & ";" & _
m_DTDERECH & ";" & m_CDPERIODICITEECH & ";" & m_MTECHEANCE1 & ";" & m_DTDEBPERECH1 & ";" & m_DTFINPERECH1 & ";" & m_MTECHEANCE2 & ";" & m_DTDEBPERECH2 & ";" & m_DTFINPERECH2 & ";" & _
m_MTECHEANCE3 & ";" & m_DTDEBPERECH3 & ";" & m_DTFINPERECH3 & ";" & m_CDTYPAMO & ";" & m_LBTYPAMO & ";" & m_TXINVPEC & ";" & m_DTDEBPERPIP & ";" & m_DTFINPERPIP & ";" & _
m_DTSAISIEPERJUSTIF & ";" & m_DTDEBPERJUSTIF & ";" & m_DTFINPERJUSTIF & ";" & m_DTDEBDERPERRGLTADA & ";" & m_DTFINDERPERRGLTADA & ";"

line = line & m_DTDERPERRGLTADA & ";" & m_MTDERPERRGLTADA & ";" & _
m_DTDEBDERPERRGLTADC & ";" & m_DTFINDERPERRGLTADC & ";" & m_DTDERPERRGLTADC & ";" & m_MTDERPERRGLTADC & ";" & m_CDSINPREPROV & ";" & m_MTTOTREGLEICIV & ";" & m_DTDEBPROV & ";" & _
m_DTFINPROV & ";" & m_INDBASREV & ";" & m_MTPREANN & ";" & m_MTCPLEST & ";" & m_MTCPLREE & ";" & m_MTPROCAL & ";" & m_MTPROIMP & ";" & m_MTPROVIT & ";" & m_MTPROPAS & ";" & m_TXPROV & ";" & _
m_TXPROPASS & ";" & m_TXTECHN & ";" & m_TXFRAIS & ";" & m_MTCAPCONORI & ";" & m_MTPREREV & ";" & m_MTPREMAJ & ";" & m_MTPANCAL & ";" & m_MTCMPPAN & ";" & m_MTCOMPAN & ";" & m_MTCOMCMP & ";" & _
m_MTPRIREG & ";" & m_MTPRIRE1 & ";" & m_MTPRIRE2 & ";" & m_CDMONNAIE & ";" & m_CDPAYS & ";" & m_CDAPPLISOURCE & ";" & m_LBCOMLIG & ";" & m_CDLOTIMPORT & ";" & m_NUTRAITP3I & ";" & _
m_NUENRP3I & ";" & m_TYPEPROVISION & ";" & m_PERALIM & ";" & m_TXLISSAGE & ";" & m_TXPROSANSLISSAGE & ";" & m_CDCATINV & ";" & m_LBCATINV & ";" & m_CDCONTENTIEUX & ";" & m_NUSINISTRE & ";" & _
m_CDCHOIXPREST & ";"

'Avant modification du champ m_FLAMORTISSABLE
'line = line & m_LBCHOIXPREST & ";" & m_MTCAPCON & ";" & m_MTCAPSSRISQ & ";" & _
'm_AGELIMINV & ";" & m_AGELIMINC & ";" & m_MTRESTEAMORPAS & ";" & m_MTAMORTPAS & ";" & _
'm_MTCOUTREFPROPAS & ";" & m_MTPROPASAPR & ";" & m_MTPROPASAVR & ";" & m_MTRESTEAMORCAL & ";" & m_MTAMORTCAL & ";" & _
'm_MTCOUTREFPROCAL & ";" & m_MTPROCALAPR & ";" & m_MTPROCALAVR & ";" & _
'm_MTINDEMRES & ";" & m_MTPROIMPAPR & ";" & m_MTPROIMPAVR & ";" & m_DTLIMPROAPR & ";" & m_TXAMORT & ";" & m_FLAMORTISSABLE & ";" & m_MTPREANN2 & ";" & m_COMMENTAIRE

'modification Ali+Alain le 15/10/2017
'line = line & m_LBCHOIXPREST & ";" & m_MTCAPCON & ";" & m_MTCAPSSRISQ & ";" & _
'm_AGELIMINV & ";" & m_AGELIMINC & ";" & m_MTRESTEAMORPAS & ";" & m_MTAMORTPAS & ";" & _
'm_MTCOUTREFPROPAS & ";" & m_MTPROPASAPR & ";" & m_MTPROPASAVR & ";" & m_MTRESTEAMORCAL & ";" & m_MTAMORTCAL & ";" & _
'm_MTCOUTREFPROCAL & ";" & m_MTPROCALAPR & ";" & m_MTPROCALAVR & ";" & _
'm_MTINDEMRES & ";" & m_MTPROIMPAPR & ";" & m_MTPROIMPAVR & ";" & m_DTLIMPROAPR & ";" & m_TXAMORT & ";" & IIf(m_FLAMORTISSABLE = True, "1", "0") & ";" & m_MTPREANN2 & ";" & m_COMMENTAIRE

'modification RS+Alain le 10/1/2018 - ajout de: IIf(IsNull(m_MTAMORTCAL), 0, m_MTAMORTCAL)
line = line & m_LBCHOIXPREST & ";" & m_MTCAPCON & ";" & m_MTCAPSSRISQ & ";" & _
m_AGELIMINV & ";" & m_AGELIMINC & ";" & m_MTRESTEAMORPAS & ";" & m_MTAMORTPAS & ";" & _
m_MTCOUTREFPROPAS & ";" & m_MTPROPASAPR & ";" & m_MTPROPASAVR & ";" & m_MTRESTEAMORCAL & ";" & IIf(IsNull(m_MTAMORTCAL), 0, m_MTAMORTCAL) & ";" & _
m_MTCOUTREFPROCAL & ";" & m_MTPROCALAPR & ";" & m_MTPROCALAVR & ";" & _
m_MTINDEMRES & ";" & m_MTPROIMPAPR & ";" & m_MTPROIMPAVR & ";" & m_DTLIMPROAPR & ";" & m_TXAMORT & ";" & IIf(m_FLAMORTISSABLE = True, "1", "0") & ";" & m_MTPREANN2 & ";" & m_COMMENTAIRE


'Problem Fields:
'8: m_MTINDEMRES
'11: m_MTCOUTREFPROCAL
'12: m_MTAMORTCAL

'###   ...= Format(dblValue,"#0.00")
  'FileNum = FreeFile
  Print #1, line
  'Print #FileNum, line
  

End Function

'---------------------------------------------------
'##ModelId=5C8A76000337
Public Function Save(DB As DataAccess) As Boolean
   Dim sql As String, ret As Boolean
   ret = True

   On Error GoTo err_Save
      
      
sql = "INSERT INTO TTPROVCOLL(CDCOMPAGNIE, CDAPPLI, CDPRODUIT, NUCONTRA, NUPRET, NUDOSSIERPREST, NUSOUSDOSSIERPREST, NUCONTRATGESTDELEG, CDGARAN, NUREFMVT, CDTYPMVT, NUTRAITE, CDCATADHESION, LBCATADHESION, CDOPTION, LBOPTIONCON, LBSOUSCR, IDASSUREAGI, IDASSURE, LBASSURE, DTNAISSASS, CDSEXASSURE, " _
                          & "IDRENTIERAGI, IDRENTIER, LBRENTIER, DTNAISSREN, IDCORENTIERAGI, IDCORENTIER, LBCORENTIER, DTNAISSCOR, DTSURVSIN, AGESURVSIN, DURPOSARR, ANCARRTRA, CDPERIODICITE, CDTYPTERME, DTEFFREN, DTLIMPROAVR, DTDERREG, DTDEBPER, DTFINPER, CDPROEVA, ANNEADHE, CDSINCON, CDMISINV, DTMISINV, NUETABLI, TXREVERSION, " _
                          & "DTCALCULPROV, DTTRAITPROV, DTCREATI, DTSIGBIA, DTDECSIN, CDRISQUE, LBRISQUE, CDSITUATSIN, DTSITUATSIN, CDPRETRATTSIN, CDCTGPRT, LBCTGPRT, DTPREECH, DTDERECH, CDPERIODICITEECH, MTECHEANCE1, DTDEBPERECH1, DTFINPERECH1, MTECHEANCE2, DTDEBPERECH2, DTFINPERECH2, MTECHEANCE3, DTDEBPERECH3, DTFINPERECH3, " _
                          & "CDTYPAMO, LBTYPAMO, TXINVPEC, DTDEBPERPIP, DTFINPERPIP, DTSAISIEPERJUSTIF, DTDEBPERJUSTIF, DTFINPERJUSTIF, DTDEBDERPERRGLTADA, DTFINDERPERRGLTADA, DTDERPERRGLTADA, MTDERPERRGLTADA, DTDEBDERPERRGLTADC, DTFINDERPERRGLTADC, DTDERPERRGLTADC, MTDERPERRGLTADC, CDSINPREPROV, MTTOTREGLEICIV, DTDEBPROV, " _
                          & "DTFINPROV, INDBASREV, MTPREANN, MTCPLEST, MTCPLREE, MTPROCAL, MTPROIMP, MTPROVIT, MTPROPAS, TXPROV, TXPROPASS, TXTECHN, TXFRAIS, MTCAPCONORI, MTPREREV, MTPREMAJ, MTPANCAL, MTCMPPAN, MTCOMPAN, MTCOMCMP, MTPRIREG, MTPRIRE1, MTPRIRE2, CDMONNAIE, CDPAYS, CDAPPLISOURCE, LBCOMLIG, CDLOTIMPORT, " _
                          & "NUTRAITP3I, NUENRP3I, TYPEPROVISION, PERALIM, TXLISSAGE, TXPROSANSLISSAGE, " _
                          & "CDCATINV, LBCATINV ,CDCONTENTIEUX, NUSINISTRE, CDCHOIXPREST, LBCHOIXPREST, MTCAPCON, MTCAPSSRISQ, " _
                          & "FLAMORTISSABLE, TXAMORT, DTLIMPROAPR, MTPROIMPAVR, MTPROIMPAPR, MTINDEMRES, MTPROCALAVR, MTPROCALAPR, MTCOUTREFPROCAL, MTAMORTCAL, MTRESTEAMORCAL, MTPROPASAVR, MTPROPASAPR, MTCOUTREFPROPAS, MTAMORTPAS, MTRESTEAMORPAS, AGELIMINC, AGELIMINV, MTPREANN2 ) "
      
 sql = sql & " VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, " _
      & "?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, " _
      & "?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, " _
      & "?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, " _
      & "?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, " _
      & "?, ?, ?, ?, ?, ?, " _
      & "?, ? ,?, ?, ?, ?, ?, ?, " _
      & "?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,? ) "


   Dim theCommand As New ADODB.Command, prm As ADODB.Parameter
   theCommand.ActiveConnection = DB.Connection
   theCommand.CommandType = adCmdText
   theCommand.CommandText = sql

   Set prm = theCommand.CreateParameter("@CDCOMPAGNIE", adNumeric, adParamInput, 0, m_CDCOMPAGNIE)
   prm.NumericScale = 0
   prm.Precision = 4
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@CDAPPLI", adNumeric, adParamInput, 0, m_CDAPPLI)
   prm.NumericScale = 0
   prm.Precision = 4
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@CDPRODUIT", adVarChar, adParamInput, IIf(IsNull(m_CDPRODUIT) Or Len(m_CDPRODUIT) = 0, 1, Len(m_CDPRODUIT)), m_CDPRODUIT)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@NUCONTRA", adVarChar, adParamInput, IIf(IsNull(m_NUCONTRA) Or Len(m_NUCONTRA) = 0, 1, Len(m_NUCONTRA)), m_NUCONTRA)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@NUPRET", adNumeric, adParamInput, 0, m_NUPRET)
   prm.NumericScale = 0
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@NUDOSSIERPREST", adVarChar, adParamInput, IIf(IsNull(m_NUDOSSIERPREST) Or Len(m_NUDOSSIERPREST) = 0, 1, Len(m_NUDOSSIERPREST)), m_NUDOSSIERPREST)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@NUSOUSDOSSIERPREST", adVarChar, adParamInput, IIf(IsNull(m_NUSOUSDOSSIERPREST) Or Len(m_NUSOUSDOSSIERPREST) = 0, 1, Len(m_NUSOUSDOSSIERPREST)), m_NUSOUSDOSSIERPREST)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@NUCONTRATGESTDELEG", adVarChar, adParamInput, IIf(IsNull(m_NUCONTRATGESTDELEG) Or Len(m_NUCONTRATGESTDELEG) = 0, 1, Len(m_NUCONTRATGESTDELEG)), m_NUCONTRATGESTDELEG)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@CDGARAN", adNumeric, adParamInput, 0, m_CDGARAN)
   prm.NumericScale = 0
   prm.Precision = 5
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@NUREFMVT", adNumeric, adParamInput, 0, m_NUREFMVT)
   prm.NumericScale = 0
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@CDTYPMVT", adNumeric, adParamInput, 0, m_CDTYPMVT)
   prm.NumericScale = 0
   prm.Precision = 4
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@NUTRAITE", adNumeric, adParamInput, 0, m_NUTRAITE)
   prm.NumericScale = 0
   prm.Precision = 4
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@CDCATADHESION", adNumeric, adParamInput, 0, m_CDCATADHESION)
   prm.NumericScale = 0
   prm.Precision = 4
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@LBCATADHESION", adVarChar, adParamInput, IIf(IsNull(m_LBCATADHESION) Or Len(m_LBCATADHESION) = 0, 1, Len(m_LBCATADHESION)), m_LBCATADHESION)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@CDOPTION", adVarChar, adParamInput, IIf(IsNull(m_CDOPTION) Or Len(m_CDOPTION) = 0, 1, Len(m_CDOPTION)), m_CDOPTION)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@LBOPTIONCON", adVarChar, adParamInput, IIf(IsNull(m_LBOPTIONCON) Or Len(m_LBOPTIONCON) = 0, 1, Len(m_LBOPTIONCON)), m_LBOPTIONCON)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@LBSOUSCR", adVarChar, adParamInput, IIf(IsNull(m_LBSOUSCR) Or Len(m_LBSOUSCR) = 0, 1, Len(m_LBSOUSCR)), m_LBSOUSCR)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@IDASSUREAGI", adVarChar, adParamInput, IIf(IsNull(m_IDASSUREAGI) Or Len(m_IDASSUREAGI) = 0, 1, Len(m_IDASSUREAGI)), m_IDASSUREAGI)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@IDASSURE", adNumeric, adParamInput, 0, m_IDASSURE)
   prm.NumericScale = 0
   prm.Precision = 10
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@LBASSURE", adVarChar, adParamInput, IIf(IsNull(m_LBASSURE) Or Len(m_LBASSURE) = 0, 1, Len(m_LBASSURE)), m_LBASSURE)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTNAISSASS", adNumeric, adParamInput, 0, m_DTNAISSASS)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@CDSEXASSURE", adVarChar, adParamInput, IIf(IsNull(m_CDSEXASSURE) Or Len(m_CDSEXASSURE) = 0, 1, Len(m_CDSEXASSURE)), m_CDSEXASSURE)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@IDRENTIERAGI", adVarChar, adParamInput, IIf(IsNull(m_IDRENTIERAGI) Or Len(m_IDRENTIERAGI) = 0, 1, Len(m_IDRENTIERAGI)), m_IDRENTIERAGI)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@IDRENTIER", adNumeric, adParamInput, 0, m_IDRENTIER)
   prm.NumericScale = 0
   prm.Precision = 10
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@LBRENTIER", adVarChar, adParamInput, IIf(IsNull(m_LBRENTIER) Or Len(m_LBRENTIER) = 0, 1, Len(m_LBRENTIER)), m_LBRENTIER)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTNAISSREN", adNumeric, adParamInput, 0, m_DTNAISSREN)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@IDCORENTIERAGI", adVarChar, adParamInput, IIf(IsNull(m_IDCORENTIERAGI) Or Len(m_IDCORENTIERAGI) = 0, 1, Len(m_IDCORENTIERAGI)), m_IDCORENTIERAGI)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@IDCORENTIER", adNumeric, adParamInput, 0, m_IDCORENTIER)
   prm.NumericScale = 0
   prm.Precision = 10
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@LBCORENTIER", adVarChar, adParamInput, IIf(IsNull(m_LBCORENTIER) Or Len(m_LBCORENTIER) = 0, 1, Len(m_LBCORENTIER)), m_LBCORENTIER)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTNAISSCOR", adNumeric, adParamInput, 0, m_DTNAISSCOR)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTSURVSIN", adNumeric, adParamInput, 0, m_DTSURVSIN)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@AGESURVSIN", adNumeric, adParamInput, 0, m_AGESURVSIN)
   prm.NumericScale = 0
   prm.Precision = 3
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DURPOSARR", adNumeric, adParamInput, 0, m_DURPOSARR)
   prm.NumericScale = 0
   prm.Precision = 3
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@ANCARRTRA", adNumeric, adParamInput, 0, m_ANCARRTRA)
   prm.NumericScale = 0
   prm.Precision = 3
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@CDPERIODICITE", adVarChar, adParamInput, IIf(IsNull(m_CDPERIODICITE) Or Len(m_CDPERIODICITE) = 0, 1, Len(m_CDPERIODICITE)), m_CDPERIODICITE)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@CDTYPTERME", adNumeric, adParamInput, 0, m_CDTYPTERME)
   prm.NumericScale = 0
   prm.Precision = 4
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTEFFREN", adNumeric, adParamInput, 0, m_DTEFFREN)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTLIMPROAVR", adNumeric, adParamInput, 0, m_DTLIMPROAVR)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTDERREG", adNumeric, adParamInput, 0, m_DTDERREG)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTDEBPER", adNumeric, adParamInput, 0, m_DTDEBPER)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTFINPER", adNumeric, adParamInput, 0, m_DTFINPER)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@CDPROEVA", adVarChar, adParamInput, IIf(IsNull(m_CDPROEVA) Or Len(m_CDPROEVA) = 0, 1, Len(m_CDPROEVA)), m_CDPROEVA)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@ANNEADHE", adNumeric, adParamInput, 0, m_ANNEADHE)
   prm.NumericScale = 0
   prm.Precision = 4
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@CDSINCON", adVarChar, adParamInput, IIf(IsNull(m_CDSINCON) Or Len(m_CDSINCON) = 0, 1, Len(m_CDSINCON)), m_CDSINCON)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@CDMISINV", adVarChar, adParamInput, IIf(IsNull(m_CDMISINV) Or Len(m_CDMISINV) = 0, 1, Len(m_CDMISINV)), m_CDMISINV)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTMISINV", adNumeric, adParamInput, 0, m_DTMISINV)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@NUETABLI", adNumeric, adParamInput, 0, m_NUETABLI)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@TXREVERSION", adNumeric, adParamInput, 0, m_TXREVERSION)
   prm.NumericScale = 2
   prm.Precision = 5
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTCALCULPROV", adNumeric, adParamInput, 0, m_DTCALCULPROV)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTTRAITPROV", adNumeric, adParamInput, 0, m_DTTRAITPROV)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTCREATI", adNumeric, adParamInput, 0, m_DTCREATI)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTSIGBIA", adNumeric, adParamInput, 0, m_DTSIGBIA)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTDECSIN", adNumeric, adParamInput, 0, m_DTDECSIN)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@CDRISQUE", adVarChar, adParamInput, IIf(IsNull(m_CDRISQUE) Or Len(m_CDRISQUE) = 0, 1, Len(m_CDRISQUE)), m_CDRISQUE)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@LBRISQUE", adVarChar, adParamInput, IIf(IsNull(m_LBRISQUE) Or Len(m_LBRISQUE) = 0, 1, Len(m_LBRISQUE)), m_LBRISQUE)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@CDSITUATSIN", adNumeric, adParamInput, 0, m_CDSITUATSIN)
   prm.NumericScale = 0
   prm.Precision = 4
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTSITUATSIN", adNumeric, adParamInput, 0, m_DTSITUATSIN)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@CDPRETRATTSIN", adVarChar, adParamInput, IIf(IsNull(m_CDPRETRATTSIN) Or Len(m_CDPRETRATTSIN) = 0, 1, Len(m_CDPRETRATTSIN)), m_CDPRETRATTSIN)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@CDCTGPRT", adVarChar, adParamInput, IIf(IsNull(m_CDCTGPRT) Or Len(m_CDCTGPRT) = 0, 1, Len(m_CDCTGPRT)), m_CDCTGPRT)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@LBCTGPRT", adVarChar, adParamInput, IIf(IsNull(m_LBCTGPRT) Or Len(m_LBCTGPRT) = 0, 1, Len(m_LBCTGPRT)), m_LBCTGPRT)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTPREECH", adNumeric, adParamInput, 0, m_DTPREECH)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTDERECH", adNumeric, adParamInput, 0, m_DTDERECH)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@CDPERIODICITEECH", adVarChar, adParamInput, IIf(IsNull(m_CDPERIODICITEECH) Or Len(m_CDPERIODICITEECH) = 0, 1, Len(m_CDPERIODICITEECH)), m_CDPERIODICITEECH)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTECHEANCE1", adNumeric, adParamInput, 0, m_MTECHEANCE1)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTDEBPERECH1", adNumeric, adParamInput, 0, m_DTDEBPERECH1)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTFINPERECH1", adNumeric, adParamInput, 0, m_DTFINPERECH1)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTECHEANCE2", adNumeric, adParamInput, 0, m_MTECHEANCE2)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTDEBPERECH2", adNumeric, adParamInput, 0, m_DTDEBPERECH2)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTFINPERECH2", adNumeric, adParamInput, 0, m_DTFINPERECH2)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTECHEANCE3", adNumeric, adParamInput, 0, m_MTECHEANCE3)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTDEBPERECH3", adNumeric, adParamInput, 0, m_DTDEBPERECH3)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTFINPERECH3", adNumeric, adParamInput, 0, m_DTFINPERECH3)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@CDTYPAMO", adVarChar, adParamInput, IIf(IsNull(m_CDTYPAMO) Or Len(m_CDTYPAMO) = 0, 1, Len(m_CDTYPAMO)), m_CDTYPAMO)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@LBTYPAMO", adVarChar, adParamInput, IIf(IsNull(m_LBTYPAMO) Or Len(m_LBTYPAMO) = 0, 1, Len(m_LBTYPAMO)), m_LBTYPAMO)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@TXINVPEC", adNumeric, adParamInput, 0, m_TXINVPEC)
   prm.NumericScale = 2
   prm.Precision = 5
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTDEBPERPIP", adNumeric, adParamInput, 0, m_DTDEBPERPIP)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTFINPERPIP", adNumeric, adParamInput, 0, m_DTFINPERPIP)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTSAISIEPERJUSTIF", adNumeric, adParamInput, 0, m_DTSAISIEPERJUSTIF)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTDEBPERJUSTIF", adNumeric, adParamInput, 0, m_DTDEBPERJUSTIF)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTFINPERJUSTIF", adNumeric, adParamInput, 0, m_DTFINPERJUSTIF)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTDEBDERPERRGLTADA", adNumeric, adParamInput, 0, m_DTDEBDERPERRGLTADA)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTFINDERPERRGLTADA", adNumeric, adParamInput, 0, m_DTFINDERPERRGLTADA)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTDERPERRGLTADA", adNumeric, adParamInput, 0, m_DTDERPERRGLTADA)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTDERPERRGLTADA", adNumeric, adParamInput, 0, m_MTDERPERRGLTADA)
   prm.NumericScale = 2
   prm.Precision = 10
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTDEBDERPERRGLTADC", adNumeric, adParamInput, 0, m_DTDEBDERPERRGLTADC)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTFINDERPERRGLTADC", adNumeric, adParamInput, 0, m_DTFINDERPERRGLTADC)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTDERPERRGLTADC", adNumeric, adParamInput, 0, m_DTDERPERRGLTADC)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTDERPERRGLTADC", adNumeric, adParamInput, 0, m_MTDERPERRGLTADC)
   prm.NumericScale = 2
   prm.Precision = 10
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@CDSINPREPROV", adVarChar, adParamInput, IIf(IsNull(m_CDSINPREPROV) Or Len(m_CDSINPREPROV) = 0, 1, Len(m_CDSINPREPROV)), m_CDSINPREPROV)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTTOTREGLEICIV", adNumeric, adParamInput, 0, m_MTTOTREGLEICIV)
   prm.NumericScale = 2
   prm.Precision = 10
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTDEBPROV", adNumeric, adParamInput, 0, m_DTDEBPROV)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTFINPROV", adNumeric, adParamInput, 0, m_DTFINPROV)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@INDBASREV", adVarChar, adParamInput, IIf(IsNull(m_INDBASREV) Or Len(m_INDBASREV) = 0, 1, Len(m_INDBASREV)), m_INDBASREV)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTPREANN", adNumeric, adParamInput, 0, m_MTPREANN)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm
   

   Set prm = theCommand.CreateParameter("@MTCPLEST", adNumeric, adParamInput, 0, m_MTCPLEST)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTCPLREE", adNumeric, adParamInput, 0, m_MTCPLREE)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTPROCAL", adNumeric, adParamInput, 0, m_MTPROCAL)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTPROIMP", adNumeric, adParamInput, 0, m_MTPROIMP)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTPROVIT", adNumeric, adParamInput, 0, m_MTPROVIT)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTPROPAS", adNumeric, adParamInput, 0, m_MTPROPAS)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@TXPROV", adNumeric, adParamInput, 0, m_TXPROV)
   prm.NumericScale = 2
   prm.Precision = 5
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@TXPROPASS", adNumeric, adParamInput, 0, m_TXPROPASS)
   prm.NumericScale = 2
   prm.Precision = 5
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@TXTECHN", adNumeric, adParamInput, 0, m_TXTECHN)
   prm.NumericScale = 2
   prm.Precision = 5
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@TXFRAIS", adNumeric, adParamInput, 0, m_TXFRAIS)
   prm.NumericScale = 2
   prm.Precision = 5
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTCAPCONORI", adNumeric, adParamInput, 0, m_MTCAPCONORI)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTPREREV", adNumeric, adParamInput, 0, m_MTPREREV)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTPREMAJ", adNumeric, adParamInput, 0, m_MTPREMAJ)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTPANCAL", adNumeric, adParamInput, 0, m_MTPANCAL)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTCMPPAN", adNumeric, adParamInput, 0, m_MTCMPPAN)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTCOMPAN", adNumeric, adParamInput, 0, m_MTCOMPAN)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTCOMCMP", adNumeric, adParamInput, 0, m_MTCOMCMP)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTPRIREG", adNumeric, adParamInput, 0, m_MTPRIREG)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTPRIRE1", adNumeric, adParamInput, 0, m_MTPRIRE1)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm
   
   Set prm = theCommand.CreateParameter("@MTPRIRE2", adNumeric, adParamInput, 0, m_MTPRIRE2)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@CDMONNAIE", adVarChar, adParamInput, IIf(IsNull(m_CDMONNAIE) Or Len(m_CDMONNAIE) = 0, 1, Len(m_CDMONNAIE)), m_CDMONNAIE)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@CDPAYS", adVarChar, adParamInput, IIf(IsNull(m_CDPAYS) Or Len(m_CDPAYS) = 0, 1, Len(m_CDPAYS)), m_CDPAYS)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@CDAPPLISOURCE", adNumeric, adParamInput, 0, m_CDAPPLISOURCE)
   prm.NumericScale = 0
   prm.Precision = 4
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@LBCOMLIG", adVarChar, adParamInput, IIf(IsNull(m_LBCOMLIG) Or Len(m_LBCOMLIG) = 0, 1, Len(m_LBCOMLIG)), m_LBCOMLIG)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@CDLOTIMPORT", adNumeric, adParamInput, 0, m_CDLOTIMPORT)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@NUTRAITP3I", adNumeric, adParamInput, 0, m_NUTRAITP3I)
   prm.NumericScale = 0
   prm.Precision = 4
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@NUENRP3I", adNumeric, adParamInput, 0, m_NUENRP3I)
   prm.NumericScale = 0
   prm.Precision = 10
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@TYPEPROVISION", adVarChar, adParamInput, IIf(IsNull(m_TYPEPROVISION) Or Len(m_TYPEPROVISION) = 0, 1, Len(m_TYPEPROVISION)), m_TYPEPROVISION)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@PERALIM", adNumeric, adParamInput, 0, m_PERALIM)
   prm.NumericScale = 0
   prm.Precision = 6
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@TXLISSAGE", adNumeric, adParamInput, 0, m_TXLISSAGE)
   prm.NumericScale = 2
   prm.Precision = 5
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@TXPROSANSLISSAGE", adNumeric, adParamInput, 0, m_TXPROSANSLISSAGE)
   prm.NumericScale = 2
   prm.Precision = 5
   theCommand.Parameters.Append prm

   
   ' Evol 2010 : Lot 2

   Set prm = theCommand.CreateParameter("@CDCATINV", adVarChar, adParamInput, IIf(IsNull(m_CDCATINV) Or Len(m_CDCATINV) = 0, 1, Len(m_CDCATINV)), m_CDCATINV)
   theCommand.Parameters.Append prm
      
   Set prm = theCommand.CreateParameter("@LBCATINV", adVarChar, adParamInput, IIf(IsNull(m_LBCATINV) Or Len(m_LBCATINV) = 0, 1, Len(m_LBCATINV)), m_LBCATINV)
   theCommand.Parameters.Append prm
      
   Set prm = theCommand.CreateParameter("@CDCONTENTIEUX", adVarChar, adParamInput, IIf(IsNull(m_CDCONTENTIEUX) Or Len(m_CDCONTENTIEUX) = 0, 1, Len(m_CDCONTENTIEUX)), m_CDCONTENTIEUX)
   theCommand.Parameters.Append prm
      
   Set prm = theCommand.CreateParameter("@NUSINISTRE", adVarChar, adParamInput, IIf(IsNull(m_NUSINISTRE) Or Len(m_NUSINISTRE) = 0, 1, Len(m_NUSINISTRE)), m_NUSINISTRE)
   theCommand.Parameters.Append prm
      
   Set prm = theCommand.CreateParameter("@CDCHOIXPREST", adVarChar, adParamInput, IIf(IsNull(m_CDCHOIXPREST) Or Len(m_CDCHOIXPREST) = 0, 1, Len(m_CDCHOIXPREST)), m_CDCHOIXPREST)
   theCommand.Parameters.Append prm
      
   Set prm = theCommand.CreateParameter("@LBCHOIXPREST", adVarChar, adParamInput, IIf(IsNull(m_LBCHOIXPREST) Or Len(m_LBCHOIXPREST) = 0, 1, Len(m_LBCHOIXPREST)), m_LBCHOIXPREST)
   theCommand.Parameters.Append prm
      
   Set prm = theCommand.CreateParameter("@MTCAPCON", adNumeric, adParamInput, 0, m_MTCAPCON)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTCAPSSRISQ", adNumeric, adParamInput, 0, m_MTCAPSSRISQ)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@FLAMORTISSABLE", adVarChar, adParamInput, 1, IIf(m_FLAMORTISSABLE = True, "1", "0"))
   theCommand.Parameters.Append prm
                          
   Set prm = theCommand.CreateParameter("@TXAMORT", adNumeric, adParamInput, 0, m_TXAMORT)
   prm.NumericScale = 2
   prm.Precision = 5
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@DTLIMPROAPR", adNumeric, adParamInput, 0, m_DTLIMPROAPR)
   prm.NumericScale = 0
   prm.Precision = 8
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTPROIMPAVR", adNumeric, adParamInput, 0, m_MTPROIMPAVR)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTPROIMPAPR", adNumeric, adParamInput, 0, m_MTPROIMPAPR)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTINDEMRES", adNumeric, adParamInput, 0, m_MTINDEMRES)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTPROCALAVR", adNumeric, adParamInput, 0, m_MTPROCALAVR)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTPROCALAPR", adNumeric, adParamInput, 0, m_MTPROCALAPR)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTCOUTREFPROCAL", adNumeric, adParamInput, 0, m_MTCOUTREFPROCAL)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTAMORTCAL", adNumeric, adParamInput, 0, m_MTAMORTCAL)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTRESTEAMORCAL", adNumeric, adParamInput, 0, m_MTRESTEAMORCAL)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTPROPASAVR", adNumeric, adParamInput, 0, m_MTPROPASAVR)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTPROPASAPR", adNumeric, adParamInput, 0, m_MTPROPASAPR)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTCOUTREFPROPAS", adNumeric, adParamInput, 0, m_MTCOUTREFPROPAS)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTAMORTPAS", adNumeric, adParamInput, 0, m_MTAMORTPAS)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@MTRESTEAMORPAS", adNumeric, adParamInput, 0, m_MTRESTEAMORPAS)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@AGELIMINC", adNumeric, adParamInput, 0, m_AGELIMINC)
   prm.NumericScale = 0
   prm.Precision = 3
   theCommand.Parameters.Append prm
   
   Set prm = theCommand.CreateParameter("@AGELIMINV", adNumeric, adParamInput, 0, m_AGELIMINV)
   prm.NumericScale = 0
   prm.Precision = 3
   theCommand.Parameters.Append prm
   
   Set prm = theCommand.CreateParameter("@MTPREANN2", adNumeric, adParamInput, 0, m_MTPREANN2)
   prm.NumericScale = 2
   prm.Precision = 15
   theCommand.Parameters.Append prm

'Debug.Print m_CDCOMPAGNIE & "," & m_CDAPPLI & "," & m_CDPRODUIT & "," & m_NUCONTRA & "," & m_NUPRET & "," & m_NUDOSSIERPREST & "," & m_NUSOUSDOSSIERPREST & "," & m_NUCONTRATGESTDELEG & "," & m_CDGARAN & "," & m_NUREFMVT & "," & m_CDTYPMVT & "," & m_NUTRAITE & "," & m_CDCATADHESION & "," & m_LBCATADHESION & "," & m_CDOPTION & "," & m_LBOPTIONCON & "," & m_LBSOUSCR & "," & m_IDASSUREAGI & "," & m_IDASSURE & "," & m_LBASSURE & "," & m_DTNAISSASS & "," & m_CDSEXASSURE & "," & m_IDRENTIERAGI & "," & m_IDRENTIER & "," & m_LBRENTIER & "," & m_DTNAISSREN & "," & m_IDCORENTIERAGI & "," & m_IDCORENTIER & "," & m_LBCORENTIER & "," & m_DTNAISSCOR & "," & m_DTSURVSIN & "," & m_AGESURVSIN & "," & m_DURPOSARR & "," & m_ANCARRTRA & "," & m_CDPERIODICITE & "," & m_CDTYPTERME & "," & m_DTEFFREN & "," & m_DTLIMPROAVR & "," & m_DTDERREG & "," & m_DTDEBPER & "," & m_DTFINPER & "," _
'& m_CDPROEVA _
'& "," & m_ANNEADHE & "," & m_CDSINCON & "," & m_CDMISINV & "," & m_DTMISINV & "," & m_NUETABLI & "," & m_TXREVERSION & "," & m_DTCALCULPROV & "," & m_DTTRAITPROV & "," & m_DTCREATI & "," & m_DTSIGBIA & "," & m_DTDECSIN & "," & m_CDRISQUE & "," & m_LBRISQUE & "," & m_CDSITUATSIN & "," & m_DTSITUATSIN & "," & m_CDPRETRATTSIN & "," & m_CDCTGPRT & "," & m_LBCTGPRT & "," & m_DTPREECH & "," & m_DTDERECH & "," & m_CDPERIODICITEECH & "," & m_MTECHEANCE1 & "," & m_DTDEBPERECH1 & "," & m_DTFINPERECH1 & "," & m_MTECHEANCE2 & "," & m_DTDEBPERECH2 & "," & m_DTFINPERECH2 & "," & m_MTECHEANCE3 & "," & m_DTDEBPERECH3 & "," & m_DTFINPERECH3 & "," & m_CDTYPAMO & "," & m_LBTYPAMO & "," & m_TXINVPEC & "," & m_DTDEBPERPIP & "," & m_DTFINPERPIP & "," & m_DTSAISIEPERJUSTIF & "," & m_DTDEBPERJUSTIF & "," & m_DTFINPERJUSTIF & "," & m_DTDEBDERPERRGLTADA & "," & m_DTFINDERPERRGLTADA _
' & "," & m_DTDERPERRGLTADA & "," & m_MTDERPERRGLTADA & "," & m_DTDEBDERPERRGLTADC & "," _
'& m_DTFINDERPERRGLTADC & "," & m_DTDERPERRGLTADC & "," & m_MTDERPERRGLTADC & "," & m_CDSINPREPROV & "," & m_MTTOTREGLEICIV & "," & m_DTDEBPROV & "," & m_DTFINPROV & "," & m_INDBASREV & "," & m_MTPREANN & "," & m_MTCPLEST & "," & m_MTCPLREE & "," & m_MTPROCAL & "," & m_MTPROIMP & "," & m_MTPROVIT & "," & m_MTPROPAS & "," & m_TXPROV & "," & m_TXPROPASS & "," & m_TXTECHN & "," & m_TXFRAIS & "," & m_MTCAPCONORI & "," & m_MTPREREV & "," & m_MTPREMAJ & "," & m_MTPANCAL & "," & m_MTCMPPAN & "," & m_MTCOMPAN & "," & m_MTCOMCMP & "," & m_MTPRIREG & "," & m_MTPRIRE1 & "," & m_MTPRIRE2 & "," & m_CDMONNAIE & "," & m_CDPAYS & "," & m_CDAPPLISOURCE & "," & m_LBCOMLIG & "," & m_CDLOTIMPORT & "," & m_NUTRAITP3I & "," & m_NUENRP3I & "," & m_TYPEPROVISION & "," & m_PERALIM _
' & "," & m_TXLISSAGE & "," & m_TXPROSANSLISSAGE & "," & m_CDCATINV & "," & m_LBCATINV & "," & m_CDCONTENTIEUX & "," & m_NUSINISTRE & "," & m_CDCHOIXPREST & "," & m_LBCHOIXPREST & "," & m_MTCAPCON & "," & m_MTCAPSSRISQ & "," & m_FLAMORTISSABLE & "," & m_TXAMORT & "," & m_DTLIMPROAPR & "," & m_MTPROIMPAVR & "," & m_MTPROIMPAPR & "," & m_MTINDEMRES & "," & m_MTPROCALAVR & "," & m_MTPROCALAPR & "," & m_MTCOUTREFPROCAL & "," & m_MTAMORTCAL & "," & m_MTRESTEAMORCAL & "," & m_MTPROPASAVR & "," & m_MTPROPASAPR & "," & m_MTCOUTREFPROPAS & "," & m_MTAMORTPAS & "," & m_MTRESTEAMORPAS & "," & m_AGELIMINC & "," & m_AGELIMINV & "," & m_MTPREANN2


' affichage des champs de TTPROVCOLL + leurs valeurs dans la fenêtre immédiate
'Debug.Print "m_CDCOMPAGNIE=" & m_CDCOMPAGNIE & "," & "m_ CDAPPLI=" & m_CDAPPLI & "," & "m_ CDPRODUIT=" & m_CDPRODUIT & "," & "m_ NUCONTRA=" & m_NUCONTRA & "," & "m_ NUPRET=" & m_NUPRET & "," & "m_ NUDOSSIERPREST=" & m_NUDOSSIERPREST _
'& "," & "m_ NUSOUSDOSSIERPREST=" & m_NUSOUSDOSSIERPREST & "," & "m_ NUCONTRATGESTDELEG=" & m_NUCONTRATGESTDELEG & "," & "m_ CDGARAN=" & m_CDGARAN & "," & "m_ NUREFMVT=" & m_NUREFMVT & "," & "m_ CDTYPMVT=" & m_CDTYPMVT & "," & _
'"m_ NUTRAITE=" & m_NUTRAITE & "," & "m_ CDCATADHESION=" & m_CDCATADHESION & "," & "m_ LBCATADHESION=" & m_LBCATADHESION & "," & "m_ CDOPTION=" & m_CDOPTION & "," & "m_ LBOPTIONCON=" & m_LBOPTIONCON & "," & "m_ LBSOUSCR=" & m_LBSOUSCR & "," & _
'"m_ IDASSUREAGI=" & m_IDASSUREAGI & "," & "m_ IDASSURE=" & m_IDASSURE & "," & "m_ LBASSURE=" & m_LBASSURE & "," & "m_ DTNAISSASS=" & m_DTNAISSASS & "," & "m_ CDSEXASSURE=" & m_CDSEXASSURE & "," & "m_ IDRENTIERAGI=" & m_IDRENTIERAGI & "," & _
'"m_ IDRENTIER=" & m_IDRENTIER & "," & "m_ LBRENTIER=" & m_LBRENTIER & "," & "m_ DTNAISSREN=" & m_DTNAISSREN & "," & "m_ IDCORENTIERAGI=" & m_IDCORENTIERAGI & "," & "m_ IDCORENTIER=" & m_IDCORENTIER & "," & "m_ LBCORENTIER=" & m_LBCORENTIER & "," & _
'"m_ DTNAISSCOR=" & m_DTNAISSCOR & "," & "m_ DTSURVSIN=" & m_DTSURVSIN & "," & "m_ AGESURVSIN=" & m_AGESURVSIN & "," & "m_ DURPOSARR=" & m_DURPOSARR & "," & "m_ ANCARRTRA=" & m_ANCARRTRA & "," & "m_ CDPERIODICITE=" & m_CDPERIODICITE & "," & _
'"m_ CDTYPTERME=" & m_CDTYPTERME & "," & "m_ DTEFFREN=" & m_DTEFFREN & "," & "m_ DTLIMPROAVR=" & m_DTLIMPROAVR & "," & "m_ DTDERREG=" & m_DTDERREG & "," & "m_ DTDEBPER=" & m_DTDEBPER & "," & "m_ DTFINPER=" & m_DTFINPER & "," & _
'"m_ CDPROEVA=" & m_CDPROEVA & "," & "m_ ANNEADHE=" & m_ANNEADHE & "," & "m_ CDSINCON=" & m_CDSINCON & "," & "m_ CDMISINV=" & m_CDMISINV & "," & "m_ DTMISINV=" & m_DTMISINV & "," & "m_ NUETABLI=" & m_NUETABLI & "," & "m_ TXREVERSION=" & _
'm_TXREVERSION & "," & "m_ DTCALCULPROV=" & m_DTCALCULPROV & "," & "m_ DTTRAITPROV=" & m_DTTRAITPROV & "," & "m_ DTCREATI=" & m_DTCREATI & "," & "m_ DTSIGBIA=" & m_DTSIGBIA & "," & "m_ DTDECSIN=" & m_DTDECSIN & "," & "m_ CDRISQUE=" & _
'm_CDRISQUE & "," & "m_ LBRISQUE=" & m_LBRISQUE & "," & "m_ CDSITUATSIN=" & m_CDSITUATSIN & "," & "m_ DTSITUATSIN=" & m_DTSITUATSIN & "," & "m_ CDPRETRATTSIN=" & m_CDPRETRATTSIN & "," & "m_ CDCTGPRT=" & m_CDCTGPRT & "," & "m_ LBCTGPRT=" & _
'm_LBCTGPRT & "," & "m_ DTPREECH=" & m_DTPREECH & "," & "m_ DTDERECH=" & m_DTDERECH & "," & "m_ CDPERIODICITEECH=" & m_CDPERIODICITEECH & "," & "m_ MTECHEANCE1=" & m_MTECHEANCE1 & "," & "m_ DTDEBPERECH1=" & m_DTDEBPERECH1 & "," & _
'"m_ DTFINPERECH1=" & m_DTFINPERECH1 & "," & "m_ MTECHEANCE2=" & m_MTECHEANCE2 & "," & "m_ DTDEBPERECH2=" & m_DTDEBPERECH2 & "," & "m_ DTFINPERECH2=" & m_DTFINPERECH2 & "," & "m_ MTECHEANCE3=" & m_MTECHEANCE3 & "," & "m_ DTDEBPERECH3=" & _
'm_DTDEBPERECH3 & "," & "m_ DTFINPERECH3=" & m_DTFINPERECH3 & "," & "m_ CDTYPAMO=" & m_CDTYPAMO & "," & "m_ LBTYPAMO=" & m_LBTYPAMO & "," & "m_ TXINVPEC=" & m_TXINVPEC & "," & "m_ DTDEBPERPIP=" & m_DTDEBPERPIP & "," & "m_ DTFINPERPIP=" & _
'm_DTFINPERPIP & "," & "m_ DTSAISIEPERJUSTIF=" & m_DTSAISIEPERJUSTIF & "," & "m_ DTDEBPERJUSTIF=" & m_DTDEBPERJUSTIF & "," & "m_ DTFINPERJUSTIF=" & m_DTFINPERJUSTIF & "," & "m_ DTDEBDERPERRGLTADA=" & m_DTDEBDERPERRGLTADA & "," & "m_ DTFINDERPERRGLTADA=" & _
'm_DTFINDERPERRGLTADA & "," & "m_ DTDERPERRGLTADA=" & m_DTDERPERRGLTADA & "," & "m_ MTDERPERRGLTADA=" & m_MTDERPERRGLTADA & "," & "m_ DTDEBDERPERRGLTADC=" & m_DTDEBDERPERRGLTADC & "," & _
'"m_ DTFINDERPERRGLTADC=" & m_DTFINDERPERRGLTADC & "," & "m_ DTDERPERRGLTADC=" & m_DTDERPERRGLTADC & "," & "m_ MTDERPERRGLTADC=" & m_MTDERPERRGLTADC & "," & "m_ CDSINPREPROV=" & m_CDSINPREPROV & "," & "m_ MTTOTREGLEICIV=" & _
'm_MTTOTREGLEICIV & "," & "m_ DTDEBPROV=" & m_DTDEBPROV & "," & "m_ DTFINPROV=" & m_DTFINPROV & "," & "m_ INDBASREV=" & m_INDBASREV & "," & "m_ MTPREANN=" & m_MTPREANN & "," & "m_ MTCPLEST=" & m_MTCPLEST & "," & "m_ MTCPLREE=" & m_MTCPLREE & "," & _
'"m_ MTPROCAL=" & m_MTPROCAL & "," & "m_ MTPROIMP=" & m_MTPROIMP & "," & "m_ MTPROVIT=" & m_MTPROVIT & "," & "m_ MTPROPAS=" & m_MTPROPAS & "," & "m_ TXPROV=" & m_TXPROV & "," & "m_ TXPROPASS=" & m_TXPROPASS & "," & "m_ TXTECHN=" & m_TXTECHN & "," & _
'"m_ TXFRAIS=" & m_TXFRAIS & "," & "m_ MTCAPCONORI=" & m_MTCAPCONORI & "," & "m_ MTPREREV=" & m_MTPREREV & "," & "m_ MTPREMAJ=" & m_MTPREMAJ & "," & "m_ MTPANCAL=" & m_MTPANCAL & "," & "m_ MTCMPPAN=" & m_MTCMPPAN & "," & "m_ MTCOMPAN=" & m_MTCOMPAN & "," & "m_ MTCOMCMP=" & m_MTCOMCMP & "," & "m_ MTPRIREG=" & m_MTPRIREG & "," & "m_ MTPRIRE1=" & m_MTPRIRE1 & "," & _
'"m_ MTPRIRE2=" & m_MTPRIRE2 & "," & "m_ CDMONNAIE=" & m_CDMONNAIE & "," & "m_ CDPAYS=" & m_CDPAYS & "," & "m_ CDAPPLISOURCE=" & m_CDAPPLISOURCE & "," & "m_ LBCOMLIG=" & m_LBCOMLIG & "," & "m_ CDLOTIMPORT=" & m_CDLOTIMPORT & "," & "m_ NUTRAITP3I=" & m_NUTRAITP3I & "," & "m_ NUENRP3I=" & m_NUENRP3I & "," & "m_ TYPEPROVISION=" & m_TYPEPROVISION & "," & _
'"m_ PERALIM=" & m_PERALIM & "," & "m_ TXLISSAGE=" & m_TXLISSAGE & "," & "m_ TXPROSANSLISSAGE=" & m_TXPROSANSLISSAGE & "," & "m_ CDCATINV=" & m_CDCATINV & "," & "m_ LBCATINV =" & m_LBCATINV & "," & "m_CDCONTENTIEUX=" & m_CDCONTENTIEUX & "," & "m_ NUSINISTRE=" & m_NUSINISTRE & "," & "m_ CDCHOIXPREST=" & m_CDCHOIXPREST & "," & "m_ LBCHOIXPREST=" & _
'm_LBCHOIXPREST & "," & "m_ MTCAPCON=" & m_MTCAPCON & "," & "m_ MTCAPSSRISQ=" & m_MTCAPSSRISQ & "," & "m_ FLAMORTISSABLE=" & m_FLAMORTISSABLE & "," & "m_ TXAMORT=" & m_TXAMORT & "," & "m_ DTLIMPROAPR=" & m_DTLIMPROAPR & "," & "m_ MTPROIMPAVR=" & _
'm_MTPROIMPAVR & "," & "m_ MTPROIMPAPR=" & m_MTPROIMPAPR & "," & "m_ MTINDEMRES=" & m_MTINDEMRES & "," & "m_ MTPROCALAVR=" & m_MTPROCALAVR & "," & "m_ MTPROCALAPR=" & m_MTPROCALAPR & "," & "m_ MTCOUTREFPROCAL=" & m_MTCOUTREFPROCAL & "," & "m_ MTAMORTCAL=" & m_MTAMORTCAL & "," & "m_ MTRESTEAMORCAL=" & m_MTRESTEAMORCAL & "," & "m_ MTPROPASAVR=" & m_MTPROPASAVR & "," & "m_ MTPROPASAPR=" & m_MTPROPASAPR & "," & "m_ MTCOUTREFPROPAS=" & m_MTCOUTREFPROPAS & "," & "m_ MTAMORTPAS=" & m_MTAMORTPAS & "," & "m_ MTRESTEAMORPAS=" & m_MTRESTEAMORPAS & "," & "m_ AGELIMINC=" & m_AGELIMINC & "," & "m_ AGELIMINV=" & m_AGELIMINV & "," & "m_ MTPREANN2 =" & m_MTPREANN2


   
   theCommand.Execute

   
   Save = ret


   Exit Function

err_Save:
   MsgBox "Erreur dans clsTTPROVCOLL::Save() : " & Err & vbLf & Err.Description, vbCritical
   Save = False
   Resume Next
End Function

