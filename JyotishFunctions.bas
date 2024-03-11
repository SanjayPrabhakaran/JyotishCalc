Attribute VB_Name = "JyotishFunctions"
'
' Swiss Ephemeris Release 1.60  9-jan-2000
'
' Declarations for Visual Basic 5.0
' The DLL file must exist in the same directory, or in a system
' directory where it can be found at runtime
'

Private Declare PtrSafe Function swe_azalt Lib "swedll64.dll" _
        ( _
          ByVal tjd_ut As Double, _
          ByVal calc_flag As Long, _
          ByRef geopos As Double, _
          ByVal atpress As Double, _
          ByVal attemp As Double, _
          ByRef xin As Double, _
          ByRef xaz As Double _
        ) As Long  'geopos must be the first of three array elements
                   'xin must be the first of two array elements
                   'xaz must be the first of three array elements

Private Declare PtrSafe Function swe_azalt_rev Lib "swedll64.dll" _
       ( _
          ByVal tjd_ut As Double, _
          ByVal calc_flag As Long, _
          ByRef geopos As Double, _
          ByRef xin As Double, _
          ByRef xout As Double _
        ) As Long  'geopos must be the first of three array elements
                   'xin must be the first of two array elements
                   'xout must be the first of three array elements

Private Declare PtrSafe Function swe_calc Lib "swedll64.dll" _
       ( _
          ByVal tjd As Double, _
          ByVal ipl As Long, _
          ByVal iflag As Long, _
          ByRef X As Double, _
          ByVal serr As String _
        ) As Long   ' x must be first of six array elements
                    ' serr must be able to hold 256 bytes

Private Declare PtrSafe Function swe_calc_d Lib "swedll64.dll" _
       ( _
          ByRef tjd As Double, _
          ByVal ipl As Long, _
          ByVal iflag As Long, _
          ByRef X As Double, _
          ByVal serr As String _
        ) As Long       ' x must be first of six array elements
                        ' serr must be able to hold 256 bytes

Private Declare PtrSafe Function swe_calc_ut Lib "swedll64.dll" _
       ( _
          ByVal tjd_ut As Double, _
          ByVal ipl As Long, _
          ByVal iflag As Long, _
          ByRef X As Double, _
          ByVal serr As String _
        ) As Long   ' x must be first of six array elements
                    ' serr must be able to hold 256 bytes

Private Declare PtrSafe Function swe_calc_ut_d Lib "swedll64.dll" _
       ( _
          ByRef tjd_ut As Double, _
          ByVal ipl As Long, _
          ByVal iflag As Long, _
          ByRef X As Double, _
          ByVal serr As String _
        ) As Long       ' x must be first of six array elements
                        ' serr must be able to hold 256 bytes

Private Declare PtrSafe Function swe_close Lib "swedll64.dll" _
       ( _
        ) As Long

Private Declare PtrSafe Function swe_close_d Lib "swedll64.dll" _
       ( _
          ByVal ivoid As Long _
        ) As Long       ' argument ivoid is ignored

Private Declare PtrSafe Sub swe_cotrans Lib "swedll64.dll" _
       ( _
          ByRef xpo As Double, _
          ByRef xpn As Double, _
          ByVal eps As Double _
        )

Private Declare PtrSafe Function swe_cotrans_d Lib "swedll64.dll" _
       ( _
          ByRef xpo As Double, _
          ByRef xpn As Double, _
          ByRef eps As Double _
        ) As Long

Private Declare PtrSafe Sub swe_cotrans_sp Lib "swedll64.dll" _
       ( _
          ByRef xpo As Double, _
          ByRef xpn As Double, _
          ByVal eps As Double _
        )

Private Declare PtrSafe Function swe_cotrans_sp_d Lib "swedll64.dll" _
       ( _
          ByRef xpo As Double, _
          ByRef xpn As Double, _
          ByRef eps As Double _
        ) As Long

Private Declare PtrSafe Sub swe_cs2degstr Lib "swedll64.dll" _
       ( _
          ByVal t As Long, _
          ByVal S As String _
        )

Private Declare PtrSafe Function swe_cs2degstr_d Lib "swedll64.dll" _
       ( _
          ByVal t As Long, _
          ByVal S As String _
        ) As Long

Private Declare PtrSafe Sub swe_cs2lonlatstr Lib "swedll64.dll" _
       ( _
          ByVal t As Long, _
          ByVal pchar As Byte, _
          ByVal mchar As Byte, _
          ByVal S As String _
        )

Private Declare PtrSafe Function swe_cs2lonlatstr_d Lib "swedll64.dll" _
       ( _
          ByVal t As Long, _
          ByRef pchar As Byte, _
          ByRef mchar As Byte, _
          ByVal S As String _
        ) As Long

Private Declare PtrSafe Sub swe_cs2timestr Lib "swedll64.dll" _
       ( _
          ByVal t As Long, _
          ByVal sep As Long, _
          ByVal supzero As Long, _
          ByVal S As String _
        )

Private Declare PtrSafe Function swe_cs2timestr_d Lib "swedll64.dll" _
       ( _
          ByVal t As Long, _
          ByVal sep As Long, _
          ByVal supzero As Long, _
          ByVal S As String _
        ) As Long

Private Declare PtrSafe Function swe_csnorm Lib "swedll64.dll" _
       ( _
          ByVal P As Long _
        ) As Long

Private Declare PtrSafe Function swe_csnorm_d Lib "swedll64.dll" _
       ( _
          ByVal P As Long _
        ) As Long

Private Declare PtrSafe Function swe_csroundsec Lib "swedll64.dll" _
       ( _
          ByVal P As Long _
        ) As Long

Private Declare PtrSafe Function swe_csroundsec_d Lib "swedll64.dll" _
       ( _
          ByVal P As Long _
        ) As Long

Private Declare PtrSafe Function swe_d2l Lib "swedll64.dll" _
       ( _
        ) As Long

Private Declare PtrSafe Function swe_d2l_d Lib "swedll64.dll" _
       ( _
        ) As Long

Private Declare PtrSafe Function swe_date_conversion Lib "swedll64.dll" _
       ( _
          ByVal Year As Long, _
          ByVal Month As Long, _
          ByVal Day As Long, _
          ByVal utime As Double, _
          ByVal cal As Byte, _
          ByRef tjd As Double _
        ) As Long

Private Declare PtrSafe Function swe_date_conversion_d Lib "swedll64.dll" _
       ( _
          ByVal Year As Long, _
          ByVal Month As Long, _
          ByVal Day As Long, _
          ByRef utime As Double, _
          ByRef cal As Byte, _
          ByRef tjd As Double _
        ) As Long

Private Declare PtrSafe Function swe_day_of_week Lib "swedll64.dll" _
       ( _
          ByVal JD As Double _
        ) As Long

Private Declare PtrSafe Function swe_day_of_week_d Lib "swedll64.dll" _
       ( _
          ByRef JD As Double _
        ) As Long

Private Declare PtrSafe Function swe_degnorm Lib "swedll64.dll" _
       ( _
          ByVal JD As Double _
        ) As Double

Private Declare PtrSafe Function swe_degnorm_d Lib "swedll64.dll" _
       ( _
          ByRef JD As Double _
        ) As Long

Private Declare PtrSafe Function swe_deltat Lib "swedll64.dll" _
       ( _
          ByVal JD As Double _
        ) As Double

Private Declare PtrSafe Function swe_deltat_d Lib "swedll64.dll" _
       ( _
          ByRef JD As Double, _
          ByRef deltat As Double _
        ) As Long

Private Declare PtrSafe Function swe_difcs2n Lib "swedll64.dll" _
       ( _
          ByVal p1 As Long, _
          ByVal p2 As Long _
        ) As Long

Private Declare PtrSafe Function swe_difcs2n_d Lib "swedll64.dll" _
       ( _
          ByVal p1 As Long, _
          ByVal p2 As Long _
        ) As Long

Private Declare PtrSafe Function swe_difcsn Lib "swedll64.dll" _
       ( _
          ByVal p1 As Long, _
          ByVal p2 As Long _
        ) As Long

Private Declare PtrSafe Function swe_difcsn_d Lib "swedll64.dll" _
       ( _
          ByVal p1 As Long, _
          ByVal p2 As Long _
        ) As Long

Private Declare PtrSafe Function swe_difdeg2n Lib "swedll64.dll" _
       ( _
          ByVal p1 As Double, _
          ByVal p2 As Double _
        ) As Double

Private Declare PtrSafe Function swe_difdeg2n_d Lib "swedll64.dll" _
       ( _
          ByRef p1 As Double, _
          ByRef p2 As Double, _
          ByRef Diff As Double _
        ) As Long

Private Declare PtrSafe Function swe_difdegn Lib "swedll64.dll" _
       ( _
          ByVal p1 As Double, _
          ByVal p2 As Double _
        ) As Long

Private Declare PtrSafe Function swe_difdegn_d Lib "swedll64.dll" _
       ( _
          ByRef p1 As Double, _
          ByRef p2 As Double, _
          ByRef Diff As Double _
        ) As Long

Private Declare PtrSafe Function swe_fixstar Lib "swedll64.dll" _
       ( _
          ByVal star As String, _
          ByVal tjd As Double, _
          ByVal iflag As Long, _
          ByRef X As Double, _
          ByVal serr As String _
        ) As Long       ' x must be first of six array elements
                        ' serr must be able to hold 256 bytes
                        ' star must be able to hold 40 bytes

Private Declare PtrSafe Function swe_fixstar_d Lib "swedll64.dll" _
       ( _
          ByVal star As String, _
          ByRef tjd As Double, _
          ByVal iflag As Long, _
          ByRef X As Double, _
          ByVal serr As String _
        ) As Long       ' x must be first of six array elements
                        ' serr must be able to hold 256 bytes
                        ' star must be able to hold 40 bytes

Private Declare PtrSafe Function swe_fixstar_ut Lib "swedll64.dll" _
       ( _
          ByVal star As String, _
          ByVal tjd_ut As Double, _
          ByVal iflag As Long, _
          ByRef X As Double, _
          ByVal serr As String _
        ) As Long       ' x must be first of six array elements
                        ' serr must be able to hold 256 bytes
                        ' star must be able to hold 40 bytes

Private Declare PtrSafe Function swe_fixstar_ut_d Lib "swedll64.dll" _
       ( _
          ByVal star As String, _
          ByRef tjd_ut As Double, _
          ByVal iflag As Long, _
          ByRef X As Double, _
          ByVal serr As String _
        ) As Long       ' x must be first of six array elements
                        ' serr must be able to hold 256 bytes
                        ' star must be able to hold 40 bytes

Private Declare PtrSafe Function swe_get_ayanamsa Lib "swedll64.dll" _
       ( _
          ByVal tjd_et As Double _
        ) As Double

Private Declare PtrSafe Function swe_get_ayanamsa_d Lib "swedll64.dll" _
       ( _
          ByRef tjd_et As Double, _
          ByRef ayan As Double _
        ) As Long

Private Declare PtrSafe Function swe_get_ayanamsa_ut Lib "swedll64.dll" _
       ( _
          ByVal tjd_ut As Double _
        ) As Double

Private Declare PtrSafe Function swe_get_ayanamsa_ut_d Lib "swedll64.dll" _
       ( _
          ByRef tjd_ut As Double, _
          ByRef ayan As Double _
        ) As Long

Private Declare PtrSafe Sub swe_get_planet_name Lib "swedll64.dll" _
       ( _
          ByVal ipl As Long, _
          ByVal pname As String _
        )

Private Declare PtrSafe Function swe_get_planet_name_d Lib "swedll64.dll" _
       ( _
          ByVal ipl As Long, _
          ByVal pname As String _
        ) As Long

Private Declare PtrSafe Function swe_get_tid_acc Lib "swedll64.dll" _
       ( _
        ) As Double

Private Declare PtrSafe Function swe_get_tid_acc_d Lib "swedll64.dll" _
       ( _
          ByRef X As Double _
        ) As Long

Private Declare PtrSafe Function swe_houses Lib "swedll64.dll" _
       ( _
          ByVal tjd_ut As Double, _
          ByVal geolat As Double, _
          ByVal geolon As Double, _
          ByVal ihsy As Long, _
          ByRef hcusps As Double, _
          ByRef ascmc As Double _
        ) As Long       ' hcusps must be first of 13 array elements
                        ' ascmc must be first of 10 array elements

Private Declare PtrSafe Function swe_houses_d Lib "swedll64.dll" _
       ( _
          ByRef tjd_ut As Double, _
          ByRef geolat As Double, _
          ByRef geolon As Double, _
          ByVal ihsy As Long, _
          ByRef hcusps As Double, _
          ByRef ascmc As Double _
        ) As Long       ' hcusps must be first of 13 array elements
                        ' ascmc must be first of 10 array elements

Private Declare PtrSafe Function swe_houses_ex Lib "swedll64.dll" _
       ( _
          ByVal tjd_ut As Double, _
          ByVal iflag As Long, _
          ByVal geolat As Double, _
          ByVal geolon As Double, _
          ByVal ihsy As Long, _
          ByRef hcusps As Double, _
          ByRef ascmc As Double _
        ) As Long       ' hcusps must be first of 13 array elements
                        ' ascmc must be first of 10 array elements

Private Declare PtrSafe Function swe_houses_ex_d Lib "swedll64.dll" _
       ( _
          ByRef tjd_ut As Double, _
          ByVal iflag As Long, _
          ByRef geolat As Double, _
          ByRef geolon As Double, _
          ByVal ihsy As Long, _
          ByRef hcusps As Double, _
          ByRef ascmc As Double _
        ) As Long       ' hcusps must be first of 13 array elements
                        ' ascmc must be first of 10 array elements

Private Declare PtrSafe Function swe_houses_armc Lib "swedll64.dll" _
       ( _
          ByVal armc As Double, _
          ByVal geolat As Double, _
          ByVal eps As Double, _
          ByVal ihsy As Long, _
          ByRef hcusps As Double, _
          ByRef ascmc As Double _
        ) As Long       ' hcusps must be first of 13 array elements
                        ' ascmc must be first of 10 array elements

Private Declare PtrSafe Function swe_houses_armc_d Lib "swedll64.dll" _
       ( _
          ByRef armc As Double, _
          ByRef geolat As Double, _
          ByRef eps As Double, _
          ByVal ihsy As Long, _
          ByRef hcusps As Double, _
          ByRef ascmc As Double _
        ) As Long       ' hcusps must be first of 13 array elements
                        ' ascmc must be first of 10 array elements

Private Declare PtrSafe Function swe_house_pos Lib "swedll64.dll" _
       ( _
          ByVal armc As Double, _
          ByVal geolat As Double, _
          ByVal eps As Double, _
          ByVal ihsy As Long, _
          ByRef xpin As Double, _
          ByVal serr As String _
        ) As Double
                        ' xpin must be first of 2 array elements

Private Declare PtrSafe Function swe_house_pos_d Lib "swedll64.dll" _
       ( _
          ByRef armc As Double, _
          ByRef geolat As Double, _
          ByRef eps As Double, _
          ByVal ihsy As Long, _
          ByRef xpin As Double, _
          ByRef hpos As Double, _
          ByVal serr As String _
        ) As Long
                        ' xpin must be first of 2 array elements

Private Declare PtrSafe Function swe_julday Lib "swedll64.dll" _
       ( _
          ByVal Year As Long, _
          ByVal Month As Long, _
          ByVal Day As Long, _
          ByVal hour As Double, _
          ByVal gregflg As Long _
        ) As Double

Private Declare PtrSafe Function swe_julday_d Lib "swedll64.dll" _
       ( _
          ByVal Year As Long, _
          ByVal Month As Long, _
          ByVal Day As Long, _
          ByRef hour As Double, _
          ByVal gregflg As Long, _
          ByRef tjd As Double _
        ) As Long

Private Declare PtrSafe Function swe_lun_eclipse_how Lib "swedll64.dll" _
       ( _
          ByVal tjd_ut As Double, _
          ByVal ifl As Long, _
          ByRef geopos As Double, _
          ByRef attr As Double, _
          ByVal serr As String _
        ) As Long

Private Declare PtrSafe Function swe_lun_eclipse_how_d Lib "swedll64.dll" _
       ( _
          ByRef tjd_ut As Double, _
          ByVal ifl As Long, _
          ByRef geopos As Double, _
          ByRef attr As Double, _
          ByVal serr As String _
        ) As Long

Private Declare PtrSafe Function swe_lun_eclipse_when Lib "swedll64.dll" _
       ( _
          ByVal tjd_start As Double, _
          ByVal ifl As Long, _
          ByVal ifltype As Long, _
          ByRef tret As Double, _
          ByVal backward As Long, _
          ByVal serr As String _
        ) As Long

Private Declare PtrSafe Function swe_lun_eclipse_when_d Lib "swedll64.dll" _
       ( _
          ByRef tjd_start As Double, _
          ByVal ifl As Long, _
          ByVal ifltype As Long, _
          ByRef tret As Double, _
          ByVal backward As Long, _
          ByVal serr As String _
        ) As Long

Private Declare PtrSafe Function swe_nod_aps Lib "swedll64.dll" _
       ( _
          ByVal tjd_et As Double, _
          ByVal ipl As Long, _
          ByVal iflag As Long, _
          ByVal method As Long, _
          ByRef xnasc As Double, _
          ByRef xndsc As Double, _
          ByRef xperi As Double, _
          ByRef xaphe As Double, _
          ByVal serr As String _
        ) As Long

Private Declare PtrSafe Function swe_nod_aps_ut Lib "swedll64.dll" _
       ( _
          ByVal tjd_ut As Double, _
          ByVal ipl As Long, _
          ByVal iflag As Long, _
          ByVal method As Long, _
          ByRef xnasc As Double, _
          ByRef xndsc As Double, _
          ByRef xperi As Double, _
          ByRef xaphe As Double, _
          ByVal serr As String _
        ) As Long

Private Declare PtrSafe Function swe_pheno Lib "swedll64.dll" _
       ( _
          ByVal tjd As Double, _
          ByVal ipl As Long, _
          ByVal iflag As Long, _
          ByRef attr As Double, _
          ByVal serr As String _
        ) As Long

Private Declare PtrSafe Function swe_pheno_ut Lib "swedll64.dll" _
       ( _
          ByVal tjd As Double, _
          ByVal ipl As Long, _
          ByVal iflag As Long, _
          ByRef attr As Double, _
          ByVal serr As String _
        ) As Long

Private Declare PtrSafe Function swe_pheno_d Lib "swedll64.dll" _
       ( _
          ByRef tjd As Double, _
          ByVal ipl As Long, _
          ByVal iflag As Long, _
          ByRef attr As Double, _
          ByVal serr As String _
        ) As Long

Private Declare PtrSafe Function swe_pheno_ut_d Lib "swedll64.dll" _
       ( _
          ByRef tjd As Double, _
          ByVal ipl As Long, _
          ByVal iflag As Long, _
          ByRef attr As Double, _
          ByVal serr As String _
        ) As Long

Private Declare PtrSafe Function swe_refrac Lib "swedll64.dll" _
       ( _
          ByVal inalt As Double, _
          ByVal atpress As Double, _
          ByVal attemp As Double, _
          ByVal calc_flag As Long _
        ) As Double

Private Declare PtrSafe Sub swe_revjul Lib "swedll64.dll" _
       ( _
          ByVal tjd As Double, _
          ByVal gregflg As Long, _
          ByRef Year As Long, _
          ByRef Month As Long, _
          ByRef Day As Long, _
          ByRef hour As Double _
        )

Private Declare PtrSafe Function swe_revjul_d Lib "swedll64.dll" _
       ( _
          ByRef tjd As Double, _
          ByVal gregflg As Long, _
          ByRef Year As Long, _
          ByRef Month As Long, _
          ByRef Day As Long, _
          ByRef hour As Double _
        ) As Long

Private Declare PtrSafe Function swe_rise_trans Lib "swedll64.dll" _
       ( _
          ByVal tjd_ut As Double, _
          ByVal ipl As Long, _
          ByVal starname As String, _
          ByVal epheflag As Long, _
          ByVal rsmi As Long, _
          ByRef geopos As Double, _
          ByVal atpress As Double, _
          ByVal attemp As Double, _
          ByRef tret As Double, _
          ByVal serr As String _
        ) As Long

Private Declare PtrSafe Sub swe_set_ephe_path Lib "swedll64.dll" _
       ( _
          ByVal path As String _
        )

Private Declare PtrSafe Function swe_set_ephe_path_d Lib "swedll64.dll" _
       ( _
          ByVal path As String _
        ) As Long

Private Declare PtrSafe Sub swe_set_jpl_file Lib "swedll64.dll" _
       ( _
          ByVal file As String _
        )

Private Declare PtrSafe Function swe_set_jpl_file_d Lib "swedll64.dll" _
       ( _
          ByVal file As String _
        ) As Long

Private Declare PtrSafe Function swe_set_sid_mode Lib "swedll64.dll" _
       ( _
          ByVal sid_mode As Long, _
          ByVal t0 As Double, _
          ByVal ayan_t0 As Double _
        ) As Long

Private Declare PtrSafe Function swe_set_sid_mode_d Lib "swedll64.dll" _
       ( _
          ByVal sid_mode As Long, _
          ByRef t0 As Double, _
          ByRef ayan_t0 As Double _
        ) As Long

Private Declare PtrSafe Function swe_set_topo Lib "swedll64.dll" _
       ( _
          ByVal geolon As Double, _
          ByVal geolat As Double, _
          ByVal altitude As Double _
        )

Private Declare PtrSafe Function swe_set_topo_d Lib "swedll64.dll" _
       ( _
          ByRef geolon As Double, _
          ByRef geolat As Double, _
          ByRef altitude As Double _
        )

Private Declare PtrSafe Sub swe_set_tid_acc Lib "swedll64.dll" _
       ( _
          ByVal X As Double _
        )

Private Declare PtrSafe Function swe_set_tid_acc_d Lib "swedll64.dll" _
       ( _
          ByRef X As Double _
        ) As Long

Private Declare PtrSafe Function swe_sidtime0 Lib "swedll64.dll" _
       ( _
          ByVal tjd_ut As Double, _
          ByVal ecl As Double, _
          ByVal nut As Double _
        ) As Double

Private Declare PtrSafe Function swe_sidtime0_d Lib "swedll64.dll" _
       ( _
          ByRef tjd_ut As Double, _
          ByRef ecl As Double, _
          ByRef nut As Double, _
          ByRef sidt As Double _
        ) As Long

Private Declare PtrSafe Function swe_sidtime Lib "swedll64.dll" _
       ( _
          ByVal tjd_ut As Double _
        ) As Double

Private Declare PtrSafe Function swe_sidtime_d Lib "swedll64.dll" _
       ( _
          ByRef tjd_ut As Double, _
          ByRef sidt As Double _
        ) As Long

Private Declare PtrSafe Function swe_sol_eclipse_how Lib "swedll64.dll" _
       ( _
          ByVal tjd_ut As Double, _
          ByVal ifl As Long, _
          ByRef geopos As Double, _
          ByRef attr As Double, _
          ByVal serr As String _
        ) As Long

Private Declare PtrSafe Function swe_sol_eclipse_how_d Lib "swedll64.dll" _
       ( _
          ByRef tjd_ut As Double, _
          ByVal ifl As Long, _
          ByRef geopos As Double, _
          ByRef attr As Double, _
          ByVal serr As String _
        ) As Long

Private Declare PtrSafe Function swe_sol_eclipse_when_glob Lib "swedll64.dll" _
       ( _
          ByVal tjd_start As Double, _
          ByVal ifl As Long, _
          ByVal ifltype As Long, _
          ByRef tret As Double, _
          ByVal backward As Long, _
          ByVal serr As String _
        ) As Long

Private Declare PtrSafe Function swe_sol_eclipse_when_glob_d Lib "swedll64.dll" _
       ( _
          ByRef tjd_start As Double, _
          ByVal ifl As Long, _
          ByVal ifltype As Long, _
          ByRef tret As Double, _
          ByVal backward As Long, _
          ByVal serr As String _
        ) As Long

Private Declare PtrSafe Function swe_sol_eclipse_when_loc Lib "swedll64.dll" _
       ( _
          ByVal tjd_start As Double, _
          ByVal ifl As Long, _
          ByRef tret As Double, _
          ByRef attr As Double, _
          ByVal backward As Long, _
          ByVal serr As String _
        ) As Long

Private Declare PtrSafe Function swe_sol_eclipse_when_loc_d Lib "swedll64.dll" _
       ( _
          ByRef tjd_start As Double, _
          ByVal ifl As Long, _
          ByRef tret As Double, _
          ByRef attr As Double, _
          ByVal backward As Long, _
          ByVal serr As String _
        ) As Long

Private Declare PtrSafe Function swe_sol_eclipse_where Lib "swedll64.dll" _
       ( _
          ByVal tjd_ut As Double, _
          ByVal ifl As Long, _
          ByRef geopos As Double, _
          ByRef attr As Double, _
          ByVal serr As String _
        ) As Long

Private Declare PtrSafe Function swe_sol_eclipse_where_d Lib "swedll64.dll" _
       ( _
          ByRef tjd_ut As Double, _
          ByVal ifl As Long, _
          ByRef geopos As Double, _
          ByRef attr As Double, _
          ByVal serr As String _
        ) As Long

Private Declare PtrSafe Function swe_time_equ Lib "swedll64.dll" _
       ( _
          ByVal tjd_ut As Double, _
          ByRef E As Double, _
          ByRef serr As String _
        ) As Long
 
' values for gregflag in swe_julday() and swe_revjul()
 Const SE_JUL_CAL As Integer = 0
 Const SE_GREG_CAL As Integer = 1

' planet and body numbers (parameter ipl) for swe_calc()
 Const SE_SUN As Integer = 0
 Const SE_MOON As Integer = 1
 Const SE_MERCURY As Integer = 2
 Const SE_VENUS As Integer = 3
 Const SE_MARS As Integer = 4
 Const SE_JUPITER As Integer = 5
 Const SE_SATURN As Integer = 6
 Const SE_URANUS As Integer = 7
 Const SE_NEPTUNE As Integer = 8
 Const SE_PLUTO   As Integer = 9
 Const SE_MEAN_NODE As Integer = 10
 Const SE_TRUE_NODE As Integer = 11
 Const SE_MEAN_APOG As Integer = 12
 Const SE_OSCU_APOG As Integer = 13
 Const SE_EARTH     As Integer = 14
 Const SE_CHIRON    As Integer = 15
 Const SE_PHOLUS    As Integer = 16
 Const SE_CERES     As Integer = 17
 Const SE_PALLAS    As Integer = 18
 Const SE_JUNO      As Integer = 19
 Const SE_VESTA     As Integer = 20
  
 Const SE_NPLANETS  As Integer = 21
 Const SE_AST_OFFSET  As Integer = 10000

' Hamburger or Uranian ficticious "planets"
 Const SE_FICT_OFFSET As Integer = 40
 Const SE_FICT_MAX  As Integer = 999 'maximum number for ficticious planets
                                     'if taken from file seorbel.txt
 Const SE_NFICT_ELEM  As Integer = 15 'number of built-in ficticious planets
 Const SE_CUPIDO As Integer = 40
 Const SE_HADES As Integer = 41
 Const SE_ZEUS As Integer = 42
 Const SE_KRONOS As Integer = 43
 Const SE_APOLLON As Integer = 44
 Const SE_ADMETOS As Integer = 45
 Const SE_VULKANUS As Integer = 46
 Const SE_POSEIDON As Integer = 47
' other ficticious bodies
 Const SE_ISIS As Integer = 48
 Const SE_NIBIRU As Integer = 49
 Const SE_HARRINGTON As Integer = 50
 Const SE_NEPTUNE_LEVERRIER As Integer = 51
 Const SE_NEPTUNE_ADAMS As Integer = 52
 Const SE_PLUTO_LOWELL As Integer = 53
 Const SE_PLUTO_PICKERING As Integer = 54

' points returned by swe_houses() and swe_houses_armc()
' in array ascmc(0...10)
 Const SE_ASC       As Integer = 0
 Const SE_MC        As Integer = 1
 Const SE_ARMC      As Integer = 2
 Const SE_VERTEX    As Integer = 3
 Const SE_EQUASC    As Integer = 4  ' "equatorial ascendant"
 Const SE_NASCMC    As Integer = 5  ' number of such points
 
' iflag values for swe_calc()/swe_calc_ut() and swe_fixstar()/swe_fixstar_ut()
Const SEFLG_JPLEPH As Long = 1
Const SEFLG_SWIEPH As Long = 2
Const SEFLG_MOSEPH As Long = 4
Const SEFLG_SPEED As Long = 256
Const SEFLG_HELCTR As Long = 8
Const SEFLG_TRUEPOS As Long = 16
Const SEFLG_J2000 As Long = 32
Const SEFLG_NONUT As Long = 64
Const SEFLG_NOGDEFL As Long = 512
Const SEFLG_NOABERR As Long = 1024
Const SEFLG_EQUATORIAL As Long = 2048
Const SEFLG_XYZ As Long = 4096
Const SEFLG_RADIANS As Long = 8192
Const SEFLG_BARYCTR As Long = 16384
Const SEFLG_TOPOCTR As Long = 32768
Const SEFLG_SIDEREAL As Long = 65536

'eclipse codes
Const SE_ECL_CENTRAL As Long = 1
Const SE_ECL_NONCENTRAL As Long = 2
Const SE_ECL_TOTAL As Long = 4
Const SE_ECL_ANNULAR As Long = 8
Const SE_ECL_PARTIAL As Long = 16
Const SE_ECL_ANNULAR_TOTAL As Long = 32
Const SE_ECL_PENUMBRAL As Long = 64
Const SE_ECL_VISIBLE As Long = 128
Const SE_ECL_MAX_VISIBLE As Long = 256
Const SE_ECL_1ST_VISIBLE As Long = 512
Const SE_ECL_2ND_VISIBLE As Long = 1024
Const SE_ECL_3RD_VISIBLE As Long = 2048
Const SE_ECL_4TH_VISIBLE As Long = 4096

'sidereal modes, for swe_set_sid_mode()
Const SE_SIDM_FAGAN_BRADLEY    As Long = 0
Const SE_SIDM_LAHIRI           As Long = 1
Const SE_SIDM_DELUCE           As Long = 2
Const SE_SIDM_RAMAN            As Long = 3
Const SE_SIDM_USHASHASHI       As Long = 4
Const SE_SIDM_KRISHNAMURTI     As Long = 5
Const SE_SIDM_DJWHAL_KHUL      As Long = 6
Const SE_SIDM_YUKTESHWAR       As Long = 7
Const SE_SIDM_JN_BHASIN        As Long = 8
Const SE_SIDM_BABYL_KUGLER1    As Long = 9
Const SE_SIDM_BABYL_KUGLER2   As Long = 10
Const SE_SIDM_BABYL_KUGLER3   As Long = 11
Const SE_SIDM_BABYL_HUBER     As Long = 12
Const SE_SIDM_BABYL_ETPSC     As Long = 13
Const SE_SIDM_ALDEBARAN_15TAU As Long = 14
Const SE_SIDM_HIPPARCHOS      As Long = 15
Const SE_SIDM_SASSANIAN       As Long = 16
Const SE_SIDM_GALCENT_0SAG    As Long = 17
Const SE_SIDM_J2000           As Long = 18
Const SE_SIDM_J1900           As Long = 19
Const SE_SIDM_B1950           As Long = 20
Const SE_SIDM_USER            As Long = 255

Const SE_NSIDM_PREDEF         As Long = 21

Const SE_SIDBITS              As Long = 256
'for projection onto ecliptic of t0
Const SE_SIDBIT_ECL_T0        As Long = 256
'for projection onto solar system plane
Const SE_SIDBIT_SSY_PLANE     As Long = 512

' modes for planetary nodes/apsides, swe_nod_aps(), swe_nod_aps_ut()
Const SE_NODBIT_MEAN        As Long = 1
Const SE_NODBIT_OSCU        As Long = 2
Const SE_NODBIT_OSCU_BAR    As Long = 3
Const SE_NODBIT_FOPOINT     As Long = 256

' indices for swe_rise_trans()
Const SE_CALC_RISE      As Long = 1
Const SE_CALC_SET       As Long = 2
Const SE_CALC_MTRANSIT      As Long = 4
Const SE_CALC_ITRANSIT      As Long = 8
Const SE_BIT_DISC_CENTER        As Long = 256 '/* to be added to SE_CALC_RISE/SET */
                    '/* if rise or set of disc center is */
                    '/* requried */
Const SE_BIT_NO_REFRACTION      As Long = 512 '/* to be added to SE_CALC_RISE/SET, */
                    '/* if refraction is not to be considered */



' bits for data conversion with swe_azalt() and swe_azalt_rev()
Const SE_ECL2HOR        As Long = 0
Const SE_EQU2HOR        As Long = 1
Const SE_HOR2ECL        As Long = 0
Const SE_HOR2EQU        As Long = 1

' for swe_refrac()
Const SE_TRUE_TO_APP        As Long = 0
Const SE_APP_TO_TRUE        As Long = 1

 
 Public Function risesetplanet( _
   lat As Double, _
   Lon As Double, _
   B As Double, _
   riseset As Long, _
   Planet As Long) _
As Double

Dim Jul_day_UT As Double, tret(10) As Double
Dim ret_flag As Double, geopos(3) As Double, serr As String
geopos(0) = Lon
geopos(1) = lat
geopos(2) = 0
    Jul_day_UT = B + 2415017.5
    ret_flag = swe_rise_trans(Jul_day_UT, Planet, "", 2, riseset, geopos(0), 1013.25, 10, tret(0), serr)
    h = tret(0) - 2415018.5
    risesetplanet = h
End Function
 
   Public Function Planet( _
   DateTime As Double, _
   PlanetID As Long) _
As Double

Dim X(6) As Double
Dim Jul_day_UT As Double
Dim i As Long, serr As String

Jul_day_UT = DateTime + 2415018.5
   i = swe_set_sid_mode(SE_SIDM_LAHIRI, 0, 0)
   i = swe_calc_ut(Jul_day_UT, PlanetID, SEFLG_SIDEREAL, X(0), serr)
Planet = X(0)

End Function
Public Function Asdt( _
   Latitude As Double, _
   Longitude As Double, _
   DateTime As Double) _
As Double

Dim Jul_day_UT As Double, X(13) As Double, A(10) As Double
Dim i As Long

   Jul_day_UT = DateTime + 2415018.5
      i = swe_houses_ex(Jul_day_UT, 65536, Latitude, Longitude, Asc("A"), X(0), A(0))
   Asdt = A(0)

End Function

Public Function Naks( _
    PlanetLongitude As Double) _
As String

Dim A(27) As String
A(0) = "Aswini"
A(1) = "Bharani"
A(2) = "Krittika"
A(3) = "Rohini"
A(4) = "Mrigashira"
A(5) = "Ardra"
A(6) = "Punarvasu"
A(7) = "Pushya"
A(8) = "Aslesha"
A(9) = "Makha"
A(10) = "Purva Phalguni"
A(11) = "Uttara Phalguni"
A(12) = "Hasta"
A(13) = "Chitra"
A(14) = "Swati"
A(15) = "Visakha"
A(16) = "Anuradha"
A(17) = "Jyestha"
A(18) = "Moola"
A(19) = "Purva Asadha"
A(20) = "Uttara Asadha"
A(21) = "Sravana"
A(22) = "Dhanistha"
A(23) = "Satabhisaj"
A(24) = "Purva Bhadrapada"
A(25) = "Uttara Bhadrapada"
A(26) = "Revati"

N = Application.WorksheetFunction.RoundUp(PlanetLongitude / 13.33, 0)

Naks = A(N - 1)

End Function

Public Function NaksPad( _
    PlanetLongitude As Double) _
As String

Dim N As Long
PlanetLongitude = PlanetLongitude / 3.333333
PlanetLongitude = Application.WorksheetFunction.RoundUp(PlanetLongitude, 1)
N = PlanetLongitude Mod 4

If N = 0 Then N = 4

NaksPad = N
MsgBox PlanetLongitude, , N

End Function

Public Function NaksDeity( _
    PlanetLongitude As Double) _
As String

Dim A(27) As String
A(0) = "Aswini-Kumar"
A(1) = "Yama"
A(2) = "Agni"
A(3) = "Brahma"
A(4) = "Chandra"
A(5) = "Rudra"
A(6) = "Aditi"
A(7) = "Brhaspati"
A(8) = "Nagas"
A(9) = "Pitaras"
A(10) = "Bhaga"
A(11) = "Aryaman"
A(12) = "Aditya"
A(13) = "Visvakarma"
A(14) = "Vayu"
A(15) = "Indragni"
A(16) = "Mitra "
A(17) = "Indra"
A(18) = "Nirriti"
A(19) = "Jal"
A(20) = "Viswadevas"
A(21) = "Vishnu"
A(22) = "Asta-Vasavas"
A(23) = "Varuna"
A(24) = "Ajaekapad"
A(25) = "Ahir-Budhnya"
A(26) = "Pushan"



N = Application.WorksheetFunction.RoundUp(PlanetLongitude / 13.33, 0)

NaksDeity = A(N - 1)

End Function


Public Function NaksLord( _
    PlanetLongitude As Double) _
As String

Dim A(27) As String
A(0) = "Ketu"
A(1) = "Venus"
A(2) = "Sun"
A(3) = "Moon"
A(4) = "Mars"
A(5) = "Rahu"
A(6) = "Jupiter"
A(7) = "Saturn"
A(8) = "Mercury"
A(9) = "Ketu"
A(10) = "Venus"
A(11) = "Sun"
A(12) = "Moon"
A(13) = "Mars"
A(14) = "Rahu"
A(15) = "Jupiter"
A(16) = "Saturn"
A(17) = "Mercury"
A(18) = "Ketu"
A(19) = "Venus"
A(20) = "Sun"
A(21) = "Moon"
A(22) = "Mars"
A(23) = "Rahu"
A(24) = "Jupiter"
A(25) = "Saturn"
A(26) = "Mercury"

N = Application.WorksheetFunction.RoundUp(PlanetLongitude / 13.33, 0)

NaksLord = A(N - 1)

End Function

Public Function Sunrise( _
   Latitude As Double, _
   Longitude As Double, _
   DateTime As Double) _
As Double

Dim Jul_day_UT As Double, tret(10) As Double
Dim ret_flag As Double, geopos(3) As Double, serr As String
geopos(0) = Longitude
geopos(1) = Latitude
geopos(2) = 0

On Error GoTo error_msg

    Jul_day_UT = DateTime + 2415017.5
    ret_flag = swe_rise_trans(Jul_day_UT, 0, "", 2, 1, geopos(0), 1013.25, 10, tret(0), serr)
    h = tret(0) - 2415018.5
    Sunrise = h
Exit Function
error_msg:
    
     Application.StatusBar = Err & ": " & Error(Err)

End Function
Public Function Sunset( _
   Latitude As Double, _
   Longitude As Double, _
   DateTime As Double) _
As Double

Dim Jul_day_UT As Double, tret(10) As Double
Dim ret_flag As Double, geopos(3) As Double, serr As String
geopos(0) = Longitude
geopos(1) = Latitude
geopos(2) = 0

    Jul_day_UT = DateTime + 2415017.5
    ret_flag = swe_rise_trans(Jul_day_UT, 0, "", 2, 2, geopos(0), 1013.25, 10, tret(0), serr)
    h = tret(0) - 2415018.5
    Sunset = h
End Function

Public Function RasiDeg( _
    PlanetLongitude As Double) _
As Double

'Finding the Longitude of a Planet in a Rasi
Dim PL As Double
PL = PlanetLongitude
PL = PL / 30
PL = PL - Application.WorksheetFunction.Floor(PL, 1)
PL = PL * 30

RasiDeg = PL

End Function

Public Function RasiNum( _
    PlanetLongitude As Double) _
As Long

'Finding the Rasi No.
Dim PL As Double
PL = PlanetLongitude
PL = PL / 30
PL = Application.WorksheetFunction.Floor(PL, 1)

RasiNum = PL + 1

End Function

Public Function HouseNum( _
    PlanetLongitude As Double, _
    AsdtLongitude As Double) _
As Long

Dim AsdtNum, PlanetNum, House As Long

AsdtNum = RasiNum(AsdtLongitude)
PlanetNum = RasiNum(PlanetLongitude)

House = PlanetNum - AsdtNum + 1

If House < 1 Then
House = House + 12
End If

HouseNum = House

End Function

Public Function DMS( _
    RasiLongitude As Double) _
As String

Dim D, M, S As Double
D = Application.WorksheetFunction.Floor(RasiLongitude, 1)
M = Application.WorksheetFunction.Floor((RasiLongitude - D) * 60, 1)
S = Application.WorksheetFunction.Round((RasiLongitude - D - M / 60) * 3600, 2)

DMS = D & ":" & M & ":" & S

End Function

Public Function Degree( _
PlanetLongitude As Double) _
As Long

Degree = Application.WorksheetFunction.Floor(RasiDeg(PlanetLongitude), 1)

End Function

Public Function Minutes( _
PlanetLongitude As Double) _
As Long

Minutes = Application.WorksheetFunction.Floor((RasiDeg(PlanetLongitude) - (Application.WorksheetFunction.Floor(RasiDeg(PlanetLongitude), 1))) * 60, 1)

End Function

Public Function Seconds( _
PlanetLongitude As Double) _
As Double

Dim D, M, S, RasiLongitude As Double

RasiLongitude = RasiDeg(PlanetLongitude)

D = Application.WorksheetFunction.Floor(RasiLongitude, 1)
M = Application.WorksheetFunction.Floor((RasiLongitude - D) * 60, 1)
S = Application.WorksheetFunction.Round((RasiLongitude - D - M / 60) * 3600, 2)

Seconds = S

End Function

' This Function finds the longitude of a planet in a sign, in a divisional chart

Public Function DivDeg( _
    RasiLongitude As Double, _
    Division As Double) _
As Double

Dim N, Div As Double

Division = 30 / Division
N = RasiLongitude / Division
Div = Application.WorksheetFunction.Floor(N, 1)
N = N - Div
N = N * 30

DivDeg = N

End Function

Public Function MFD( _
SignIndex As Long) _
As String

Dim Sign(11) As String
Sign(0) = "M"
Sign(1) = "F"
Sign(2) = "D"
Sign(3) = "M"
Sign(4) = "F"
Sign(5) = "D"
Sign(6) = "M"
Sign(7) = "F"
Sign(8) = "D"
Sign(9) = "M"
Sign(10) = "F"
Sign(11) = "D"

MFD = Sign(SignIndex - 1)

End Function

Public Function Oddity( _
SignIndex As Long) _
As String

Dim Sign(11) As String

Sign(0) = "O"
Sign(1) = "E"
Sign(2) = "O"
Sign(3) = "E"
Sign(4) = "O"
Sign(5) = "E"
Sign(6) = "O"
Sign(7) = "E"
Sign(8) = "O"
Sign(9) = "E"
Sign(10) = "O"
Sign(11) = "E"

Oddity = Sign(SignIndex - 1)

End Function

Public Function Elements( _
SignIndex As Long) _
As String

Dim Sign(11) As String

Sign(0) = "F"
Sign(1) = "E"
Sign(2) = "A"
Sign(3) = "W"
Sign(4) = "F"
Sign(5) = "E"
Sign(6) = "A"
Sign(7) = "W"
Sign(8) = "F"
Sign(9) = "E"
Sign(10) = "A"
Sign(11) = "W"

Elements = Sign(SignIndex - 1)

End Function

Public Function SolarLunarHalf( _
SignIndex As Long) _
As String

Dim Sign(11) As String

Sign(0) = "L"
Sign(1) = "L"
Sign(2) = "L"
Sign(3) = "L"
Sign(4) = "S"
Sign(5) = "S"
Sign(6) = "S"
Sign(7) = "S"
Sign(8) = "S"
Sign(9) = "S"
Sign(10) = "L"
Sign(11) = "L"

SolarLunarHalf = Sign(SignIndex - 1)


End Function


' This function finds the Longitude of the planet in the divisional
' Charts. Although it is known that the rasi chart is the chart, where
' we have the longitudes of the planets, we can still map the longitude
' of the planets in each division. This is done by equating 30degs of
' the rasi chart with the n no. of divisions in the divisional charts,
' and the longitude if found out finding position of the planets from
' the beginning of the zodiac. i.e., 0 Deg Aries.

' After the longitude of a planet is found in the divisional chart, we
' can do all the processing, which we normally do with the longitude
' of planets in the rasi chart.


Public Function DivPlanetLongitude( _
PlanetLongitude As Double, _
Division As Long, _
Var As Long) _
As Double

Select Case Division

'--------------------
Case 1
DivPlanetLongitude = PlanetLongitude

'--------------------
Case 2
    Select Case Var
    
    Case 1
    'Hora based on the the planet to be placed in the 1st/7th, based on the planet in
        '1st/ 2nd Hora
        
        Dim Hora1, StartPoint21 As Long
        Hora1 = Application.WorksheetFunction.RoundUp(RasiDeg(PlanetLongitude) / 15, 0)

            If Hora1 = 1 Then
                StartPoint21 = 0
                ElseIf Hora1 = 2 Then
                StartPoint21 = 180
            End If

        DivPlanetLongitude = DegRnd(StartPoint21 + RasiNum(PlanetLongitude) * 30 - 30 + DivDeg(RasiDeg(PlanetLongitude), 2))

    
    Case 2
    ' Hora based on Standard Parashara's Interpretation
        Dim Hora21 As Long, Hora22 As String, Startpoint22 As Double
    
        Hora21 = Application.WorksheetFunction.RoundUp(RasiDeg(PlanetLongitude) / 15, 0)
        Hora22 = Oddity(RasiNum(PlanetLongitude))
        
            If Hora21 = 1 And Hora22 = "O" Then
            Startpoint22 = 120
            ElseIf Hora21 = 2 And Hora22 = "O" Then
            Startpoint22 = 90
            ElseIf Hora21 = 1 And Hora22 = "E" Then
            Startpoint22 = 90
            ElseIf Hora21 = 2 And Hora22 = "E" Then
            Startpoint22 = 120
            End If
        
        DivPlanetLongitude = DegRnd(Startpoint22 + DivDeg(RasiDeg(PlanetLongitude), 2))
    
    Case 3
     ' PD Hora
        Dim Hora3 As Long, Startpoint23 As Double
    
        Hora3 = Application.WorksheetFunction.RoundUp(RasiDeg(PlanetLongitude) / 15, 0)
        
        Startpoint23 = (RasiNum(PlanetLongitude) - 1) * 2 * 30 + (Hora3 - 1) * 30
        
        DivPlanetLongitude = DegRnd(Startpoint23 + DivDeg(RasiDeg(PlanetLongitude), 2))
    
    Case 4
      ' Jagannath Hora
      ' A planet placed in the solar sign in a solar hora or in the lunar signs
      ' in a lunar hora continues in the same sign in hora chart else it is placed
      ' in the seventh sign
        
        Dim Hora41 As Long, Hora42, Hora43, SMHora As String, Startpoint242 As Double
    
        Hora41 = Application.WorksheetFunction.RoundUp(RasiDeg(PlanetLongitude) / 15, 0)
        Hora42 = Oddity(RasiNum(PlanetLongitude))
        Hora43 = SolarLunarHalf(RasiNum(PlanetLongitude))
        
            If Hora41 = 1 And Hora42 = "O" Then
            SMHora = "S"
            ElseIf Hora41 = 2 And Hora42 = "O" Then
            SMHora = "M"
            ElseIf Hora41 = 1 And Hora42 = "E" Then
            SMHora = "M"
            ElseIf Hora41 = 2 And Hora42 = "E" Then
            SMHora = "S"
            End If
           
            If SMHora = "S" And Hora43 = "L" Then
            Startpoint242 = 180
            ElseIf SMHora = "S" And Hora43 = "S" Then
            Startpoint242 = 0
            ElseIf SMHora = "M" And Hora43 = "L" Then
            Startpoint242 = 0
            ElseIf SMHora = "M" And Hora43 = "S" Then
            Startpoint242 = 180
            End If
            
        
        DivPlanetLongitude = DegRnd(Startpoint242 + RasiNum(PlanetLongitude) * 30 - 30 + DivDeg(RasiDeg(PlanetLongitude), 2))

        Case 5
      ' Kashinath Hora
      ' The real hora chart variation that shows wealth, as taught in the tradition of
      ' Sri Achyuta Dasa, is known as "Kashinatha Hora" chart, named after Pt. Kashinath Rath.
      ' This chart is based on the classicfication of signs into signs that are strong during
      ' the day and signs that are strong during the night.
      
      
      ' Sun represents day-strong signs, i.e. Leo, Virgo, Libra, Scorpio, Aquarius and Pisces.
      ' Moon represents night-strong signs, i.e. Aries, Taurus, Gemini, Cancer, Sagittarius
      ' and Capricorn.
 
      ' It may be noted that each planet owns exactly two signs owns one day-strong sign and
      ' one night-strong signs. There is a fable that says that Sun and Moon originally
      ' owned six signs each and gave away one sign each from their six signs to Mercury,
      ' Venus, Mars, Jupiter and Saturn. Thus, Sun's hora and Moon's hora refer to the
      ' day-strong and night-strong signs among the two signs owned by a planet.
      ' For this purpose, Leo and Cancer form a pair, even though they are not owned by the
      ' same planet.
 
        Dim Hora51, Hora52, Startpoint25(11, 1) As Long, N As Double
    
        Hora51 = Application.WorksheetFunction.RoundUp(RasiDeg(PlanetLongitude) / 15, 0)
        Hora52 = RasiNum(PlanetLongitude)
        
        Startpoint25(0, 0) = 210
        Startpoint25(0, 1) = 0
        Startpoint25(1, 0) = 30
        Startpoint25(1, 1) = 180
        Startpoint25(2, 0) = 150
        Startpoint25(2, 1) = 60
        Startpoint25(3, 0) = 90
        Startpoint25(3, 1) = 120
        Startpoint25(4, 0) = 120
        Startpoint25(4, 1) = 90
        Startpoint25(5, 0) = 60
        Startpoint25(5, 1) = 150
        Startpoint25(6, 0) = 180
        Startpoint25(6, 1) = 30
        Startpoint25(7, 0) = 0
        Startpoint25(7, 1) = 210
        Startpoint25(8, 0) = 330
        Startpoint25(8, 1) = 240
        Startpoint25(9, 0) = 270
        Startpoint25(9, 1) = 300
        Startpoint25(10, 0) = 300
        Startpoint25(10, 1) = 270
        Startpoint25(11, 0) = 240
        Startpoint25(11, 1) = 330
        
        N = Startpoint25(Hora52 - 1, Hora51 - 1)
        
        DivPlanetLongitude = DegRnd(N + DivDeg(RasiDeg(PlanetLongitude), 2))
    
    
    End Select
    

'--------------------
Case 3
' Drekkana Chart
    Select Case Var
    
    Case 1
    'Parasara Drekkana
        Dim Drek1, StartPoint31 As Long
        Drek = Application.WorksheetFunction.RoundUp(RasiDeg(PlanetLongitude) / 10, 0)
                
                If Drek1 = 1 Then
                    StartPoint31 = 0
                    ElseIf Drek1 = 2 Then
                    StartPoint31 = 120
                    ElseIf Drek1 = 3 Then
                    StartPoint31 = 240
                End If
                
        DivPlanetLongitude = DegRnd(StartPoint31 + RasiNum(PlanetLongitude) * 30 - 30 + DivDeg(RasiDeg(PlanetLongitude), 3))
    
    Case 2
    'Jagannath Drekkna
        Dim Drek21, StartPoint32 As Long, drek22 As String
        Drek21 = Application.WorksheetFunction.RoundUp(RasiDeg(PlanetLongitude) / 10, 0)
        drek22 = MFD(RasiNum(PlanetLongitude))
                
                If drek22 = "M" And Drek21 = 1 Then
                StartPoint32 = 0
                ElseIf drek22 = "M" And Drek21 = 2 Then
                StartPoint32 = 120
                ElseIf drek22 = "M" And Drek21 = 3 Then
                StartPoint32 = 240
                ElseIf drek22 = "F" And Drek21 = 1 Then
                StartPoint32 = 240
                ElseIf drek22 = "F" And Drek21 = 2 Then
                StartPoint32 = 0
                ElseIf drek22 = "F" And Drek21 = 3 Then
                StartPoint32 = 120
                ElseIf drek22 = "D" And Drek21 = 1 Then
                StartPoint32 = 120
                ElseIf drek22 = "D" And Drek21 = 2 Then
                StartPoint32 = 240
                ElseIf drek22 = "D" And Drek21 = 3 Then
                StartPoint32 = 0
                End If
                
        DivPlanetLongitude = DegRnd(StartPoint32 + RasiNum(PlanetLongitude) * 30 - 30 + DivDeg(RasiDeg(PlanetLongitude), 3))
    
    Case 3
    ' PT Drekkana
        Dim Drek3 As Long, Startpoint33 As Double
    
        Drek3 = Application.WorksheetFunction.RoundUp(RasiDeg(PlanetLongitude) / 10, 0)
        
        Startpoint33 = (RasiNum(PlanetLongitude) - 1) * 3 * 30 + (Drek3 - 1) * 30
        
        DivPlanetLongitude = DegRnd(Startpoint33 + DivDeg(RasiDeg(PlanetLongitude), 3))
    
    
    Case 4
    ' Somnath Drekkana
        Dim Drek41, Drek42, StartPoint34 As Long, Drek43 As String
        Drek41 = Application.WorksheetFunction.RoundUp(RasiDeg(PlanetLongitude) / 10, 0)
                
        Drek42 = RasiNum(PlanetLongitude)
        Drek43 = Oddity(RasiNum(PlanetLongitude))
        
        If Drek42 = 1 Then
        StartPoint34 = 0
        ElseIf Drek42 = 2 Then
        StartPoint34 = 360
        ElseIf Drek42 = 3 Then
        StartPoint34 = 90
        ElseIf Drek42 = 4 Then
        StartPoint34 = 270
        ElseIf Drek42 = 5 Then
        StartPoint34 = 180
        ElseIf Drek42 = 6 Then
        StartPoint34 = 180
        ElseIf Drek42 = 7 Then
        StartPoint34 = 270
        ElseIf Drek42 = 8 Then
        StartPoint34 = 90
        ElseIf Drek42 = 9 Then
        StartPoint34 = 0
        ElseIf Drek42 = 10 Then
        StartPoint34 = 360
        ElseIf Drek42 = 11 Then
        StartPoint34 = 90
        ElseIf Drek42 = 12 Then
        StartPoint34 = 270
        End If
        
                
        If Drek43 = "O" Then
        DivPlanetLongitude = StartPoint34 + RasiDeg(PlanetLongitude) * 3
        
        ElseIf Drek43 = "E" Then
        DivPlanetLongitude = DegRnd(StartPoint34 - RasiDeg(PlanetLongitude) * 3)
        
        End If
        
    End Select
    
       
'--------------------
Case 4
Select Case Var
    Case 1
    
    ' Parashara Chaturthamsa Chart
    
      Dim Chatur, StartPoint41 As Long
        Chatur = Application.WorksheetFunction.RoundUp(RasiDeg(PlanetLongitude) / 7.5, 0)
                
                If Chatur = 1 Then
                    StartPoint41 = 0
                    ElseIf Chatur = 2 Then
                    StartPoint41 = 90
                    ElseIf Chatur = 3 Then
                    StartPoint41 = 180
                    ElseIf Chatur = 4 Then
                    StartPoint41 = 270
                End If
                
        DivPlanetLongitude = DegRnd(StartPoint41 + RasiNum(PlanetLongitude) * 30 - 30 + DivDeg(RasiDeg(PlanetLongitude), 4))
    
 
    
    Case 2
    'Cyclical Chaturthamsa Chart
    
     Dim Chatur2 As Long, Startpoint42 As Double
    
     Chatur2 = Application.WorksheetFunction.RoundUp(RasiDeg(PlanetLongitude) / 7.5, 0)
        
     Startpoint42 = (RasiNum(PlanetLongitude) - 1) * 4 * 30 + (Chatur2 - 1) * 30
       
     DivPlanetLongitude = DegRnd(Startpoint42 + DivDeg(RasiDeg(PlanetLongitude), 3))
    
   End Select

'--------------------
Case 5
' Panchamsa Chart

Select Case Var

    Case 1
    ' Cyclical Panchamsa Chart
    Dim Panch1 As Long, Startpoint51 As Double
     
    Panch1 = Application.WorksheetFunction.RoundUp(RasiDeg(PlanetLongitude) / 6, 0)
           
    Startpoint51 = (RasiNum(PlanetLongitude) - 1) * 5 * 30 + (Panch1 - 1) * 30
           
    DivPlanetLongitude = DegRnd(Startpoint51 + DivDeg(RasiDeg(PlanetLongitude), 3))

End Select

'--------------------
Case 6
' Shasthamsa Chart
    
    Dim Shast As String, Startpoint6 As Double
        
    Shast = Oddity(RasiNum(PlanetLongitude))
            
        If Shast = "O" Then
            Startpoint6 = 0
            ElseIf Shast = "E" Then
            Startpoint6 = 180
        End If
            
    DivPlanetLongitude = DegRnd(Startpoint6 + RasiDeg(PlanetLongitude) * 6)


'--------------------
Case 7
' Saptamsa Chart

    Dim Sapt As String, Startpoint7 As Double
    
        Sapt = Oddity(RasiNum(PlanetLongitude))
        
            If Sapt = "O" Then
                Startpoint7 = 0
                ElseIf Sapt = "E" Then
                Startpoint7 = 180
            End If
        
    DivPlanetLongitude = DegRnd((RasiNum(PlanetLongitude) - 1) * 30 + Startpoint7 + RasiDeg(PlanetLongitude) * 7)


'--------------------
Case 8
' Asthamsa Chart


'--------------------
Case 9
' Navamsa Chart

Dim M, Startpoint9 As Double

M = MFD(RasiNum(PlanetLongitude))

    If M = "M" Then
       Startpoint9 = 0
       ElseIf M = "F" Then
       Startpoint9 = 240
       ElseIf M = "D" Then
       Startpoint9 = 120
    End If

DivPlanetLongitude = DegRnd((RasiNum(PlanetLongitude) - 1) * 30 + Startpoint9 + RasiDeg(PlanetLongitude) * 9)

'--------------------
Case 10
' Dasamsa Chart

  Dim Dasm As String, Startpoint10 As Double
    
        Dasm = Oddity(RasiNum(PlanetLongitude))
        
            If Dasm = "O" Then
               Startpoint10 = 0
            ElseIf Dasm = "E" Then
               Startpoint10 = 240
               
            End If
        
    DivPlanetLongitude = DegRnd((RasiNum(PlanetLongitude) - 1) * 30 + Startpoint10 + RasiDeg(PlanetLongitude) * 10)

'--------------------
Case 11
' Rudramsa Chart
' The Rudramsas' are always reckoned zodiacally and every sign has eleven
' Rudramsas (Ekadasamsas) of 2o43'38" arc. For Aries these are reckoned from Aries;
' for Taurus they are reckoned Pisces; fro Gemini; reckoned from Aquarius; for
' Cancer reckoned from Capricorn; for Leo reckoned from Sagittarius and so on

    Dim Rudr1, Rudr2, StartPoint11 As Long
        Rudr1 = Application.WorksheetFunction.RoundUp(RasiDeg(PlanetLongitude) / 30 * 11, 0)
        Rudr2 = RasiNum(PlanetLongitude)
        
        
        StartPoint11 = 360 - (Rudr2 - 1) * 30
                       
        DivPlanetLongitude = DegRnd(StartPoint11 + RasiDeg(PlanetLongitude) * 11)

'--------------------
Case 12
' Dwadasamsa Chart

   DivPlanetLongitude = DegRnd((RasiNum(PlanetLongitude) - 1) * 30 + RasiDeg(PlanetLongitude) * 12)
   

'--------------------
Case 16
' Shodasamsa Chart

Dim Shod, Startpoint16 As Double

Shod = MFD(RasiNum(PlanetLongitude))

    If Shod = "M" Then
       Startpoint16 = 0
       ElseIf Shod = "F" Then
       Startpoint16 = 120
       ElseIf Shod = "D" Then
       Startpoint16 = 240
    End If

DivPlanetLongitude = DegRnd(Startpoint16 + RasiDeg(PlanetLongitude) * 16)


'--------------------
Case 20
' Vimsamsa Chart

Dim Vims, Startpoint20 As Double

Vims = MFD(RasiNum(PlanetLongitude))

    If Vims = "M" Then
       Startpoint20 = 0
       ElseIf Vims = "F" Then
       Startpoint20 = 240
       ElseIf Vims = "D" Then
       Startpoint20 = 120
    End If

DivPlanetLongitude = DegRnd(Startpoint20 + RasiDeg(PlanetLongitude) * 20)


'--------------------
Case 24
' Siddhamsa Chart

Dim Sidd As String, Startpoint24 As Double
    
Sidd = Oddity(RasiNum(PlanetLongitude))
        
    If Sidd = "O" Then
        Startpoint24 = 120
        ElseIf Sidd = "E" Then
        Startpoint24 = 90
    End If
        
DivPlanetLongitude = DegRnd(Startpoint24 + RasiDeg(PlanetLongitude) * 24)


'--------------------
Case 27
' Nakshatramsa Chart

Dim Naks, Startpoint27 As Double

Naks = Elements(RasiNum(PlanetLongitude))

    If Naks = "F" Then
       Startpoint27 = 0
       ElseIf Naks = "E" Then
       Startpoint27 = 90
       ElseIf Naks = "A" Then
       Startpoint27 = 180
       ElseIf Naks = "W" Then
       Startpoint27 = 270
    End If

DivPlanetLongitude = DegRnd(Startpoint27 + RasiDeg(PlanetLongitude) * 27)

'--------------------
Case 30
' Tattva based Trimsamsa Chart

Select Case Var
    
    Case 1
    
    Dim Trims1, StartPoint30 As Long, Trims2 As Double
    Trims = Oddity(RasiNum(PlanetLongitude))
    Trims2 = RasiDeg(PlanetLongitude)
    
    
    If Trims = "O" Then
        If Trims2 > 0 And Trims2 <= 5 Then
        StartPoint30 = 0
        ElseIf Trims2 > 5 And Trims2 <= 10 Then
        StartPoint30 = 300
        ElseIf Trims2 > 10 And Trims2 <= 18 Then
        StartPoint30 = 240
        ElseIf Trims2 > 18 And Trims2 <= 25 Then
        StartPoint30 = 60
        ElseIf Trims2 > 25 And Trims2 <= 30 Then
        StartPoint30 = 180
        End If
    ElseIf Trims = "E" Then
        If Trims2 > 30 And Trims2 <= 5 Then
        StartPoint30 = 0
        ElseIf Trims2 > 5 And Trims2 <= 12 Then
        StartPoint30 = 150
        ElseIf Trims2 > 12 And Trims2 <= 20 Then
        StartPoint30 = 330
        ElseIf Trims2 > 20 And Trims2 <= 25 Then
        StartPoint30 = 270
        ElseIf Trims2 > 25 And Trims2 <= 30 Then
        StartPoint30 = 210
        End If
    End If
    
    DivPlanetLongitude = DegRnd(StartPoint30 + DivDeg(RasiDeg(PlanetLongitude), 30))
    
    Case 2
    ' Privritti Trimsa Trimsamsa
        Dim Trimsa1 As Long, Startpoint302 As Double
    
        Trimsa1 = Application.WorksheetFunction.RoundUp(RasiDeg(PlanetLongitude) / 10, 0)
        
        Startpoint302 = (RasiNum(PlanetLongitude) - 1) * 30 * 30 + (Trimsa1 - 1) * 30
        
        DivPlanetLongitude = DegRnd(Startpoint302 + DivDeg(RasiDeg(PlanetLongitude), 30))
    
    Case 3
    ' Trimsamsa- Like Shastyamsa
       DivPlanetLongitude = DegRnd((RasiNum(PlanetLongitude) - 1) * 30 + RasiDeg(PlanetLongitude) * 30)

End Select
        
'--------------------
Case 40
' Kshavedamsa Chart

Dim Khav As String, Startpoint40 As Double
    
Khav = Oddity(RasiNum(PlanetLongitude))
        
    If Khav = "O" Then
        Startpoint40 = 0
        ElseIf Khav = "E" Then
        Startpoint40 = 180
    End If
        
DivPlanetLongitude = DegRnd(Startpoint40 + RasiDeg(PlanetLongitude) * 40)


'--------------------
Case 45
' Akshavedamsa Chart

Dim Aksh, Startpoint45 As Double

Aksh = MFD(RasiNum(PlanetLongitude))

    If Aksh = "M" Then
       Startpoint45 = 0
       ElseIf Aksh = "F" Then
       Startpoint45 = 120
       ElseIf Aksh = "D" Then
       Startpoint45 = 240
    End If

DivPlanetLongitude = DegRnd(Startpoint45 + RasiDeg(PlanetLongitude) * 45)


'--------------------
Case 60
' Shastyamsa Chart

   DivPlanetLongitude = DegRnd((RasiNum(PlanetLongitude) - 1) * 30 + RasiDeg(PlanetLongitude) * 60)


'--------------------
Case 108
' Ashttottaramsa Chart


'--------------------
Case 144
' Dwadasamsa- Dwadamsasa Chart


End Select

End Function

' The following function finds the modulus from any division in
' decimals, which is not done by normal MOD function of Excel

Public Function Mod2( _
Num As Double, _
Div As Double) _
As Double

Dim D As Double
D = Num / Div
D = D - Application.WorksheetFunction.Floor(D, 1)
Mod2 = D * Div

End Function

' This function rounds up the degree in between 0 and 360

Public Function DegRnd( _
Degree As Double) _
As Double

If Degree < 0 Then
    Do While Degree < 0
    Degree = Degree + 360
    Loop
    DegRnd = Degree
    
ElseIf Degree > 0 Then
    DegRnd = Mod2(Degree, 360)

End If

End Function

' This function gives the sign name based on the longitude of
' a body in the zodiac

Public Function Rasi( _
PlanetLongitude As Double) _
As String

Dim N As Long, R(11) As String

R(0) = "Ar"
R(1) = "Ta"
R(2) = "Ge"
R(3) = "Cn"
R(4) = "Le"
R(5) = "Vi"
R(6) = "Li"
R(7) = "Sc"
R(8) = "Sg"
R(9) = "Cp"
R(10) = "Aq"
R(11) = "Pi"

N = Application.WorksheetFunction.RoundUp(PlanetLongitude / 30, 0)
Rasi = R(N - 1)

End Function


' This function shall find the Transit of a planet
' after a given date on a given longitude

Public Function PlanetTransit( _
Degree As Double, _
DateTime As Double, _
PlanetID As Long) _
As Date

'Defining the speed of various planets
Const Sun = 1
Const Moon = 2.5
Const Mars = 1
Const Merc = 1
Const Jup = 0.833
Const Venus = 1
Const Sat = 0.033
Const Rahu = 0.055
Const Ketu = 0.055

Dim Longt, Diff, Days As Double

Longt = Planet(DateTime, PlanetID)

Do Until Longt = Degree
    If Degree < Longt Then
    Diff = Degree - Longt + 360
    Else
    Diff = Degree - Longt
    End If

DateTime = DateTime + 0.8 * Diff

Longt = Planet(DateTime, PlanetID)
Loop

PlanetTransit = CDate(DateTime)

End Function

' This function put the planets in the zodiac in their
' respective rasis in a divisional chart. This function
' can not handle composite divisional chart


Public Function RasiContent( _
Rasi As Long, _
DT As Double, _
Division As Long, _
Var As Long, _
AsdtLongitude As Double) _
As String

Dim PlanetID As Long
Dim R As Long
Dim Z As String
Dim P(10) As String

P(0) = "Sun"
P(1) = "Moon"
P(2) = "Merc"
P(3) = "Venus"
P(4) = "Mars"
P(5) = "Jupiter"
P(6) = "Saturn"
P(7) = ""
P(8) = ""
P(9) = ""
P(10) = "Rahu"

For PlanetID = 0 To 10
  
R = RasiNum(DivPlanetLongitude(Planet(DT, PlanetID), Division, Var))

If R = Rasi Then
Z = Z & " " & P(PlanetID)
End If

Next PlanetID

R = RasiNum(DivPlanetLongitude(AsdtLongitude, Division, Var))
If R = Rasi Then
Z = Z & " " & "Asdt"
End If

R = RasiNum(DivPlanetLongitude(DegRnd(Planet(DT, 10) + 180), Division, Var))
If R = Rasi Then
Z = Z & " " & "Ketu"
End If

RasiContent = Z

End Function

Public Function DivName( _
Division As Long, _
Var As Long) _
As String

Dim N(144, 6) As String

N(1, 1) = "Rasi"
N(2, 1) = "Hora (1/7 House)"
N(2, 2) = "Parashara Hora"
N(2, 3) = "Parivrittidvaya  Hora"
N(2, 4) = "Jagannath Hora"
N(2, 5) = "Kashinath Hora"
N(3, 1) = "Parashara Drekkana"
N(3, 2) = "Jagannath Drekkana"
N(3, 3) = "Parivrittitraya  Drekkana"
N(3, 4) = "Somnath Drekkana"
N(4, 1) = "Parashara Chaturthamsa"
N(4, 2) = "Cyclical Chaturthamsa"
N(5, 1) = "Cyclical Panchamsa"
N(6, 1) = "Shasthamsa"
N(7, 1) = "Saptamsa"
N(8, 1) = "Ashtamsa"
N(9, 1) = "Navamsa"
N(10, 1) = "Dasamsa"
N(11, 1) = "Rudramsa"
N(12, 1) = "Dwadasamsa"
N(16, 1) = "Shodasamsa"
N(20, 1) = "Vimsamsa"
N(24, 1) = "Siddhamsa"
N(27, 1) = "Nakshatramsa"
N(30, 1) = "Standard Trimsamsa"
N(30, 2) = "Parivritti Trimsa Trimsamsa"
N(30, 3) = "Trimsamsa (Like Shastyamsa)"
N(40, 1) = "Khavedamsa"
N(45, 1) = "Akshavedamsa"
N(60, 1) = "Shastyamsa"
N(108, 1) = "Astottaramsa"
N(144, 1) = "Dwadasamsa- Dwadasamsa"

DivName = N(Division, Var)

End Function

Sub EditThis()
'
' Macro1 Macro
'
'
End Sub



