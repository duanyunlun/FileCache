'10Aug09(DNY) - ProTrace Trip Manifest Create,View,Edit
'
'rev.19Oct09(DNY) - v1.1 add support for dest-63 YRCReimer-Buffalo
'rev.17Nov09(DNY) - add RDWY Southbound Pkup Tor (dest-69)
'                 - add SameDay (68-68) manifest with auto-correspondence (see SavMnfst)
'rev.22Mar10(DNY) - v1.2 special handling for origin-QuikX-to-NPME southbound shipments
'rev.30Aug11(DNY) - add 'Save - No Print' (mnfst sent to dispatch) feature '//^\\_//^\\
'rev.05Oct11(DNY) - Z-Status (delivery status) assignment for drop trailers, access controlled by inits
'    .    .    .    .    .    .    .    .    .    .    .    .    .    .    .    .    .    .    .    .    .
'    .    .    .    .    .    .    .    .    .    .    .    .    .    .    .    .    .    .    .    .    .
'vars for controlling grid movement & paired-rows display
'  tricky coding - a pair of thick & thin rows are worked as if they are one
'                - thick row contains probill, thin row contains shipment delivery attributes/warnings/etc. in italics
'                - complicated by 2 factors: 1) thin row uses merged cells which causes the grid to stop
'                                               automatic row hiliting - so its all done manually
'                                            2) MS Bug! Up Arrow Key is processed by grid before seen by form control!!
'                                                       while the Down Arrow Key can be manually trapped!!
'
'rev.08Dec11(DNY) - 2-sided DR Print using common routines also in Stack Manifest & ProTrace
'                 - single shipment & whole manifest
'
'rev.30Jan12(DNY) - support for agent Maritime Ontario (81), southbound customs partner YU Express (55)
'
'rev.10Feb12(DNY) - food/drug shipment trailer seal/lock accnt code processing & auto-email
'
'rev.29May12(DNY) - 'SDT' owner-prefix changed to 'SP' - see sdt*
'
'rev.05Nov12(DNY) - convert from QXTI to NDTL as Western Can. partner - Linehaul code 52
'
'rev.15Apr13(DNY) - Kingsway Quebec manifest beyond-pronumbers capture - Linehaul code 40
'
'rev.27May13(DNY) - add 'Storage Trailer' manifest support (orig/dest code = 31-31), create 'S' status history entry
'rev.29May        - update PrintDR routine to support 'LOAD ON STORAGE' routing message
'
'rev.23Aug13(DNY) - add domestic LineHaul Arrivals -> new status code '^', updates stathist & timeline for all manifested pros
'
'rev.04Oct13(DNY) - add Guilbault Transport (TGBT) agent transfer support - trigger for EDI 204 sending
'
'rev.28Oct13(DNY) - add Clarke Transport (SCAC CDTT) agnt transfer support for Western Can. intermodal
'
'rev.27Nov13(DNY) - re-vamp of printing manifests & gate pass - new status code - 'Loaded' - link Gate Pass to On-Dlvy status!!
'
'rev.19Feb14(DNY) - phase 1 integration/triggering of TrailerTrace events/locations updates
'
'rev 04Apr14(DNY) - auto-un-post All-Short shipments if manifested & 'dispatched' (gate pass print) - location = manifest trailer
'
'rev 17Apr14(DNY) - TrailerTrace DDE link on trip number, link to trigger gate pass print + On-Dlvy statuses
'
'rev 30May14(DNY) - FastFrate TOR scac=FAST, Linehaul code 95 FFTR
'rev 10Jun14(DNY) - FastFrate 10-digit beyond pro capture
'rev 07Jul14(DNY) - FastFrate MTL scac=FAST, Linehaul code 96 FFMT
'
'rev.12Sep14(DNY) - Meyers Transport transfers (replacing KTown/Kingston code 8 for Kingston area deliveries)
'                   scac=MEYT, Linehaul code 89
'
'rev.14Oct14(DNY) - Sherway Warehouse standing appointment - eMail appointment advisory & BOL's on close trailer (Loading status)
'
'rev.07Apr15(DNY) - WARD southbound w/beyond pros
'
'rev.13Jul15(DNY) - Bourret ex-Tor (BTLM) shipment handling (re-consign, auto3rdparty, etc.)
'
'rev.14Oct15(DNY) - Strip Manifest launch
'
'rev.09Nov15(DNY) - SEFL southbound
'
'rev.18Mar16(DNY) - RDWY-BUF southbound (linehaul code 55)
'
'rev.02Jun16(DNY) - EXLA-BUF southbound (linehaul code 54)
'
'rev.19Jul16(DNY) - Schaeffler Can. south to Schaeffler U.S. - alert eMail closed to load (ACE manifest generated)
'
'rev.14Sep16(DNY) - Can. Tire hi-vis remark on all shipments as consignee re: waiting times
'
'rev.27Sep16(DNY) - drop dest codes 52,66,90 chg 48,64,65
'

'equipment service & type/use advisory vars
Dim clstc As Control    'source trailer or unit listbox control - link to pop-up menus
Dim cownc As String     'owner of trailer or unit
Dim cscac As String     'owner SCAC
Dim intv As Integer     'see tmr
Private Type svctyp  'service variable type
  t As String   'equipment eg.trailer
  o As String   'owner scac
  c As String   'service code eg. PM,AN,RR,OS for trailers
End Type
Dim sv() As svctyp 'form-level dynamic array
Dim lts_t As String, lts_d As String 'last-saved trailer status trailer, datetime
Dim ltm_t As String, ltm_d As String 'last_saved trailer e-mail trailer, datetime

Dim mef As Boolean      'rights to new 'Save & Dispatch option' '//^\\_//^\\
Dim oprnt0 As Boolean, oprnt1 As Boolean, oprnt2 As Boolean '21Nov13 replace o_prnt(0,1,2) as save/print selectors
Dim oprntL As Boolean, oprntG As Boolean '21Nov13
Dim g2t As Integer      'current grid top row
Dim oldrow As Integer   'previous grid row       (thick row of thick-thin pair)
Dim newrow As Integer   'new or current grid row    '    '   '   '     '    '
Dim curow As Integer    'current grid row
Dim g2foc As Boolean    'True when grid 'g2' has focus & is linked to MouseWheel API calls
Dim prono As Boolean    'recursion control for g2pro_Change event
Dim notunfr As Boolean  '  "         "
Dim nottrfr As Boolean  '  "         "
Dim nothtrfr As Boolean '  "         "
Dim notttr As Boolean   '  "         "
Dim nottun As Boolean   '  "         "
Dim loading As Boolean  'True when loading an existing manifest
Dim savprnt As Boolean  'True when save triggered by Print request
Dim g2profoc As Boolean 'True when forcing focus to pro entry field
Dim noprotrp As Boolean 'True when printing gate pass only (no pro trip)
Dim chksvcno As Boolean 'True if block check trailer service codes (for DDE link with TrailerTrace)
Dim lhcod(1 To 99) As String   'linehaul codes for trip#
Dim destcty(1 To 99) As String 'destination city code
Dim prolst(1 To 104, 0 To 1) As String 'original pro list for existing manifest, 0 = pro#, 1 = status (remove, edit)
Dim uownr(0 To 99, 0 To 1) As String   'outside tractor owner code/name
Dim tsc(0 To 19, 0 To 2) As String     'trailer owner code, name, scac
'''Dim townr(0 To 99, 0 To 1) As String   'outside trailer owner code/name
Dim holtr(0 To 9999, 0 To 1) As String 'holland trailers
Dim uholtr As Boolean                  'True of holland currently in display list (lst_htr)
Dim exltr(0 To 30000, 0 To 2) As String
Dim uexltr As Boolean
''Dim npmtr(0 To 2000, 0 To 2) As String 'newpenn trailers
''Dim unpmtr As Boolean                  'True if newpenn currently in display list
''Dim odftr(0 To 20, 0 To 2) As String   'ODF city trailers
''Dim uodftr As Boolean                  'true if ODF currently in display list
Dim dockpkup As Boolean                'True if dock-pickup manifest
Dim prnttom As Boolean                 'True if save initiated by "Print Gate Pass" button-click
Dim rettrip As String                  'return trip# for border linehaul trips
Dim cnewclick As Boolean               'True if 'Create New/Edit Trip' button was clicked to trigger a manifest save

Dim pch As Boolean    'True = skip pro_Change subroutine
Dim orgch As Boolean  ' "   "  "   orig   "      "
Dim dstch As Boolean  ' "   "  "   dest   "      "
Dim tnw As Boolean, ted As Boolean 'new & edit mode indicators
Dim eddte As Date     'delivery date of trip in edit mode (allows date change)
Dim epro(1 To 2, 1 To 104) As String   'edit pro buffer - 1,x=pro 2,x=action (remove or update with edit trip info)

Dim ucmdl As Boolean
Dim remoteproed As String
Dim remoteproedb As Boolean

Dim emconn As Boolean, emevent As Boolean 'indicators for e-mail server TCP conversation
Dim emdata As String
Dim forcetrailerlistrefresh As Boolean
'Dim ef As Object      'mailer log file

Dim gpcnt As Integer   'count of gate passes printed in a session

Dim defprn As String   'client default Win printer
Dim odefprn As Object  'client default printer object
Dim gateprn As String     'dedicated gate pass printer
Dim ogateprn As Object    'gate pass printer object
Dim inbondprn As String   'dedicated inland-inbond DR printer
Dim oinbondprn As Object  'inland-inbond DR printer object
Dim prngatpas As Boolean
Dim prninbond As Boolean
Dim fo As Object       'file I/O
Private obj As PictPlus60Vic.clsPicturePlus 'image viewer object for manifest printing

'AutoMail emailer vars
Dim aemfrm As String, aemto As String, aemcc As String, aembcc As String, aempoll As String
Dim aemsubj As String, aembody As String, aminipth As String, aempth As String, umail As String

Dim logorec As Integer 'DR print - 3rd Party company logo image record no. - if configured as attribute
Dim logox As Double, logoy As Double, logow As Double, logoh As Double
Dim pspdy As Boolean
Dim cpypp As Integer   'DR print - copies per page
Dim dr_tel As String
Dim edibill As Boolean 'DR print from partner EDI if True (else from Speedy probill)
Dim logotype As String  'jpg, gif - upper-left logo #@#
Dim wlogotype As String 'wgif, wjpg - watermark #@#

'-DRPrint----maintain identical between all apps with DR Print ------------------------------------------------------------------------------------
Dim pgc As Integer, pgn As Integer '03Nov11 print vars: totpgs, pgno
Dim pgi() As String  'array of additional info strings
Dim barc0w As Long, barc0h As Long 'base size of barc(0) picture box used for various barcodes & images
Dim pb As New ADODB.Recordset
Dim drscac As String, transn As String, brk As String, cdte As String, ctim As String, drcsano As String
Dim entport As String, cport As String, notes As String, apptstr As String, drrow As String, drrowdct As String
Dim edipro As String, drsdir As String 'partner(advance)pro, path to ProTrace dir logo images
Dim drtrailer As String, drrun As String, drinport As String, drdoor As String, reprn As String, loadid As String
Dim bt(0 To 5) As String   'billto/3rd party details
Dim reqdeldte As String, opcltim As String, indte As String 'service date, rcvng hrs on service date, declared-inland date
Dim drpo As String, drbyndscac As String, drdlvyterm As String, drrowdc As String, drrowman As String 'PO# string, dlvyterm='Dest' box
Dim opcl(1 To 5) As String, tel1(0 To 1) As String, tel2(0 To 1) As String 'rcvng hrs for week, 12Nov13 contact tel#'s for ship(0) and cons(1)
Dim drprnting As Boolean, rcvnghrs As Boolean, frzbl As Boolean
Dim accs(1 To 19, 0 To 1) As String 'cons/shipr account codes: xx,0 = manual, xx,1=preset
Dim acccnt As Integer, reprnorigcnt As Integer, reprnwhsecnt As Integer, accs2 As String 'accs array counter, reprint counters, holiday hrs 2nd-line
Dim csa4 As Boolean, drinbond As Boolean, cbsac As Boolean, csac As Boolean, apptreqd As Boolean
Dim prnodfl As Boolean, usnorthbnd As Boolean, western As Boolean, atlantic As Boolean, ontque As Boolean
Dim prnward As Boolean, prnbtlm As Boolean, prnsefl As Boolean, prnexla As Boolean, prnhmes As Boolean, prnpyle As Boolean
Dim pndx As Long
Dim apptdate As Date
Dim homehardware34henry As Boolean
'-DRPrint----------------------------------------------------------------------------------------------------------------------------------------
Dim drmnfst As Boolean

'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'API functions to support asynchronous Run cmd
Private Const infinit = &HFFFF 'give full windows timeout length to process
Private Const synch = &H100000 'synchonous mode = 'pause this app until external process completes'
Private Declare Sub WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long)
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDA As Long, ByVal bIH As Integer, ByVal dwPID As Long) As Long
Private Declare Sub CloseHandle Lib "kernel32.dll" (ByVal hObject As Long)
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'

Private Sub c_canchtr_Click()
 frhtr.Visible = False: fr1.Enabled = True
 If t_tr.Visible Then t_tr.SetFocus
End Sub

Private Sub c_canctr_Click()
 frtr.Visible = False: fr1.Enabled = True
 If t_tr.Visible Then t_tr.SetFocus
End Sub

Private Sub c_cancun_Click()
 frun.Visible = False: fr1.Enabled = True
 If t_un.Visible Then t_un.SetFocus
End Sub

Private Sub c_clr_Click()
'' MsgBox GetNxtISARecno
'' MsgBox GetNxtSTRecno
 ClrFrm
End Sub

Private Function GetTermDisp$(t%)
 Dim dm As String
 dm = ""
 Select Case t
   Case 1: dm = "TOR"
   Case 3: dm = "MIS"
   Case 4: dm = "MIL"
   Case 5: dm = "LON"
   Case 6: dm = "WIN"
   Case 8: dm = "PIC"
   Case 10: dm = "BRO"
   Case 11: dm = "LAC" '"VAU  in 2022
   Case 12: dm = "MTL"
   Case 25: dm = "VAN"
   Case 26: dm = "EDM"
   Case 27: dm = "CGY"
   Case 28: dm = "WPG"
   Case 34: dm = "TOR"
   Case 35: dm = "LAC"
   Case 85: dm = "MCN"
   Case 86: dm = "DAR"
   Case 91: dm = "DRU"
 End Select
 GetTermDisp = dm
End Function

Private Sub Disabl4Prn() 'disable controls while printing
 frw.Visible = True: Refresh: DoEvents: MousePointer = 11: fr1.Enabled = False: frg1.Enabled = False
 c_new.Enabled = False: c_prnt.Enabled = False: c_clr.Enabled = False: c_fgo.Enabled = False
End Sub

Private Sub EnablPrn() 're-enable controls after printing
 frw.Visible = False: MousePointer = 0: fr1.Enabled = True: frg1.Enabled = True: c_new.Enabled = True
 c_prnt.Enabled = True: c_clr.Enabled = True: c_fgo.Enabled = True
End Sub

Private Function GetPrnCtrl$() 'get next unique print control#
 On Error Resume Next
 cmd.CommandText = "update tripcntr set prntctrl = last_insert_id(prntctrl + 1)": Err = 0: b = False: cmd.Execute
 If Err <> 0 Then
   MsgBox "FAIL Generate Next Print Control Number!" & vbCrLf & vbCrLf & Err.Description, , "Manifest"
   GetPrnCtrl = "Z" & CStr(Month(Now)) & CStr(Day(Now)) & CStr(Hour(Now)) & CStr(Minute(Now)): Exit Function
 End If
 rs.Open "select last_insert_id()" 'complete mySQL method that ensures only this user gets this next-number
 If rs.EOF Or Err <> 0 Then
   rs.Close: MsgBox "FAIL Retrieve Next-Number Print Control Index!" & vbCrLf & vbCrLf & Err.Description, , "Manifest"
   GetPrnCtrl = "Z" & CStr(Month(Now)) & CStr(Day(Now)) & CStr(Hour(Now)) & CStr(Minute(Now)): Exit Function
 End If
 GetPrnCtrl = rs(0): rs.Close 'next-number trip# retrieved
End Function

Private Sub ClrFrm()
 Dim j As Integer
 savprnt = False: g2profoc = False: frtom.Visible = False: g2.Visible = False
 trip = "": trailer = "": unit = "": t_uno = "": t_tro = "SP": nottun = True: t_un = ""
 nottun = False: notttr = True: t_tr = "": notttr = False: trlrapp = "": driver = "": rtrpn = "": l_loadby = ""
 seal = "": ldsht = "": descr = "": l_or = "": l_de = "": orig = "": dest = "": trpno = "": l_arr(0) = "": l_arr(1) = ""
 assgn = "": ch_htr = False: g2pro = "": c_prntbols.Visible = False: c_arr.Visible = False '21Nov13 pfr.Visible = False
 c_prntdr.Visible = False: oprnt0 = False: oprnt1 = False: oprnt2 = False: oprntL = False: oprntG = False
 For j = 0 To 206 Step 2
   If g2.TextMatrix(j, 17) <> "" Or g2.TextMatrix(j, 2) <> "" Then Clrg2 j
 Next j
 g2.TopRow = 0
 If oldrow > -1 Then HiRo oldrow, 0, 0
 HiRo 0, 1, 0
 oldrow = 0: g2pro.BackColor = &HD6E6E6
 If l_ne = "EDIT" Then
   mdte = Now: mdte_Change
 End If
 g2.Visible = True: twt = "": tsh = "": tcu = "": tcucnt = "": rettrip = "": l_ne = "NEW": l_ne.ForeColor = &H70905C
 c_new.Caption = "Create NEW Trip": c_new.BackColor = &HC0D0C0: c_ldtrp.Caption = "Load by Trip#"
 c_ldtrp.BackColor = &HC9CAC9: c_lddte.Caption = "Load Trip by Date-Trailer"
 c_lddte.BackColor = &HC9CAC9: c_zstat.BackColor = &HD0D8DC
End Sub

Private Sub Clrg2(x%)
 Dim i As Integer, j As Integer
 For i = x To x + 1
   For j = 1 To 17: g2.TextMatrix(i, j) = "": Next j
   g2.TextMatrix(i, 20) = "": g2.TextMatrix(i, 21) = "": g2.TextMatrix(i, 22) = "": g2.TextMatrix(i, 23) = ""
 Next i
 g2.TextMatrix(x + 1, 1) = "" '{@} new last freight location field
 For j = 0 To 2: g2.TextMatrix(x + 1, j) = " ": Next j
 For j = 3 To 9 '{@} was 1 to 9
   g2.TextMatrix(x + 1, j) = " " 'put space in cells to force cell-merge
 Next j
End Sub

Private Function mChkDig$(p$) 'Meyers variation of Mod10 check-digit algorithm
 Dim j As Integer, k As Single
 Dim s(1 To 9) As Variant
 k = 0
 For j = 1 To 9 Step 2: s(j) = Mid(p, j, 1): Next j 'buff odds
 For j = 2 To 8 Step 2: s(j) = Mid(p, j, 1) * 2: Next j 'buff 2 x evens
 For j = 1 To 9 'sum up buff values where if any buff value > 10 then sum its 2 digits together (eg. 18 becomes 1+8=9)
   If s(j) < 10 Then k = k + s(j) Else k = k + Left(s(j), 1) + Right(s(j), 1)
 Next j
 mChkDig = Right$(Format$(k / 10#, "#.0"), 1)
End Function


Private Sub c_emptyrpt_Click()
 emptyfrm.Show 1
End Sub

Private Sub c_loc_Click()
 dockchk.Show 1
End Sub

Private Sub c_clrpros_Click()
 retrpfrm.Show 1
End Sub

Private Sub c_fgo_Click()
   '' switched from last entered/edited to tripnum to get arround 'arrivals' re-date-timing an older trip and hiding a pre-set city trip  21Aug20(DNY)
      ''rs.Open "select tripnum from trippro where pro='" & t_fpro & "' order by date desc, cast(hr as unsigned) desc, cast(min as unsigned) desc limit 1"
   rs.Open "select tripnum from trippro where pro='" & t_fpro & "' order by tripnum desc limit 1"
   If rs.EOF Then
     rs.Close: MsgBox "Manifest Not Found for Pro: " & t_fpro, , "Manifest": Exit Sub
   End If
   
 'End If
 ClrFrm
 trip = rs(0): rs.Close
 c_ldtrp_Click
End Sub


Private Sub c_lddte_Click() 'find trip by date-trailer
 Dim j As Integer, k As Integer
 Dim dm As String, dm2 As String, sqs As String, trap As String
 Dim sp() As String
 Dim stort As Boolean
 
 If c_ldtrp.BackColor = &HB0BAE0 Then 'exit edit mode
   ClrFrm
   mdte = Now: mdte_Change
   Exit Sub
 End If
  
 dm = Trim$(t_tr)
 If dm = "" Then
   MsgBox "Missing Trailer Entry!", , "Manifest"
   Exit Sub
 End If
 
 'fix case where full speedy trailer description is auto-found
 If InStr(t_tr, " ") > 0 Then 'if spaces in trailer entry ..
   sp = Split(t_tr, " "): dm = sp(UBound(sp))  'split by space, take last segment only
 End If
  
 'look for city terminal trailer prefix, if found, process trailer# as-is
 Select Case Left$(t_tr, 3)
   Case "BRO", "LON", "MTL", "PIC", "TOR", "WIN": GoTo STLDDTE
 End Select
 'look for std. SP trailer prefix - process as-is
 If Left$(t_tr, 2) = "SP" Then
   lst_tro.ListIndex = 0: lst_tro.Selected(0) = True: t_tro = "SP": GoTo STLDDTE
 End If
 'look for owner-SCAC prefixed trailer# - process as-is
 If IsNumeric(Left(t_tr, 4)) = False Then
   For j = 0 To lst_tro.ListCount - 1
     sp = Split(lst_tro.List(j), Chr(9)): k = Len(sp(0))
     If Left$(t_tr, k) = Trim$(sp(0)) Then
       lst_tro.ListIndex = j: lst_tro.Selected(j) = True: t_tro = sp(0): GoTo STLDDTE
     End If
   Next j
 End If
  
 Select Case t_tro 'owner
   Case ""    'Speedy or owner added as prefix directly to trailer#
     If IsNumeric(Left(dm, 1)) = True Then
       'no prefix, assume Speedy trailer
       Select Case Left$(dm, 3)
         Case "03C", "05C", "05M" 'old UCan trailers
         Case Else
           Do While Left$(dm, 1) = "0" 'remove leading zeros
             dm = Mid$(dm, 2)
           Loop
       End Select
     Else
       'has owner prefix - process as-is - fall-thru
     End If
   Case "SP", "SDT" 'Speedy 'sdt*
     Select Case Left$(dm, 3)
       Case "03C", "05C", "05M"
       Case Else
         Do While Left$(dm, 1) = "0" 'remove leading zeros
           dm = Mid$(dm, 2)
         Loop
     End Select
   Case Else 'other owner
     j = Len(t_tro)
     If Left$(dm, j) = t_tro Then
       'prefix matches owner code, process trailer entry as-is
     Else
       'non-match - add prefix to trailer entry
       dm = t_tro & dm
     End If
 End Select
  
STLDDTE:
  
 stort = False
 If orig = "31" Or dest = "31" Then 'if storage trailer origin/dest code then look for previous use of trailer as storage within 6 months of mnfst date
   sqs = "select date,unit,seal,origin,dest,loader,carrier,trailer,assign,run,tripnum,unitpre,trlrpre,driver,returntripnum, " & _
         "cast(arrdatetime as char),arrinit from trip where date > '" & _
          Format$(DateValue(mdte) - 180, "YYYY-MM-DD") & "' and trailer='" & dm & "' and origin='31' and dest='31' order by tripnum desc limit 1"
   stort = True: GoTo CLDDTEQ
 End If
  
 '             0    1    2     3     4     5       6       7       8    9    10      11      12       13        14
 '                 15               16       17
 sqs = "select date,unit,seal,origin,dest,loader,carrier,trailer,assign,run,tripnum,unitpre,trlrpre,driver,returntripnum, " & _
       "cast(arrdatetime as char),arrinit,loadedby from trip where date='" & _
        Format$(DateValue(mdte), "YYYY-MM-DD") & "' and trailer='" & dm & "'"
 If Trim$(trlrapp) = "" Then 'no run specified
   'fall-thru without specifying run#
 Else
   sqs = sqs & " and run='" & CStr(Val(trlrapp)) & "'"
 End If
 sqs = sqs & " order by tripnum desc limit 1"
CLDDTEQ:
 rs.Open sqs
 If rs.EOF Then
   rs.Close
   If stort Then
     sqs = "Trailer " & trailer & " NOT Used for Storage Manifest in 6-Month Period: " & Format$(DateValue(mdte) - 180, "DDD DD-MMM-YYYY") & "  to  " & _
            Format$(DateValue(mdte), "DDD DD-MMM-YYYY") & vbCrLf & vbCrLf & "To search further, set 'Date of Trip' back 6 or more months."
     MsgBox sqs, , "Manifest": Exit Sub
   End If
   If t_tro = "" And IsNumeric(Trim$(t_tr)) = True Then
      sqs = "select date,unit,seal,origin,dest,loader,carrier,trailer,assign,run,tripnum,unitpre,trlrpre,driver,returntripnum, " & _
            "cast(arrdatetime as char),arrinit,loadedby from trip where date='" & _
             Format$(DateValue(mdte), "YYYY-MM-DD") & "' and trailer like '%" & Trim$(t_tr) & "'"
      If Trim$(trlrapp) <> "" Then sqs = sqs & " and run='" & CStr(Val(trlrapp)) & "'"
      sqs = sqs & " order by tripnum desc limit 1"
      rs.Open sqs
      If Not rs.EOF Then GoTo FMNFST
      rs.Close
   End If
   If Val(trlrapp) > 0 Then dm = dm & " Run " & CStr(Val(trlrapp))
   MsgBox "Trip Not Found for Trailer " & dm & " on " & Format$(DateValue(mdte), "DD-MMM-YY"), , "Manifest"
   Exit Sub
 End If

FMNFST:
'count matching records
 k = 0
 Do
   k = k + 1: rs.MoveNext
 Loop Until rs.EOF
 rs.MoveFirst
 
 If k > 1 Then 'multiple matches, display to user for selection
   gen1 = "-99"
   entcnt.Show 1
   If gen1 = "-99" Then Exit Sub
   rs.Close
   '               0    1    2     3     4    5       6       7      8     9    10      11      12      13       14
   '                 15               16       17
   sqs = "select date,unit,seal,origin,dest,loader,carrier,trailer,assign,run,tripnum,unitpre,trlrpre,driver,returntripnum, " & _
         "cast(arrdatetime as char),arrinit,loadedby from trip where tripnum='" & gen1 & "'"
   rs.Open sqs
 End If
 
 frw.Visible = True: Refresh: DoEvents
 DispEditMode
 frw.Visible = False
 
End Sub

Private Sub c_ldtrp_Click() 'load existing trip by trip#
 Dim j As Integer
 Dim dm As String
 
 If l_ne = "EDIT" Then
   ClrFrm
   mdte = Now: mdte_Change: Exit Sub
 End If
 
 dm = Trim$(trip)
 Select Case Len(dm)
   Case 6
     If IsNumeric(Left(dm, 1)) Then dm = "tripnum='" & dm
   Case 7: dm = "assign='" & dm
   Case Else: Exit Sub
 End Select
 
 frw.Visible = True: frw.Refresh
 If rs.State <> 0 Then rs.Close
 '                 0    1    2     3     4     5      6       7       8    9    10      11      12      13         14
 '                  15                16       17
 rs.Open "select date,unit,seal,origin,dest,loader,carrier,trailer,assign,run,tripnum,unitpre,trlrpre,driver,returntripnum," & _
         "cast(arrdatetime as char),arrinit,loadedby from trip where " & dm & "' order by date desc limit 1"
 If rs.EOF Then
   rs.Close
   MsgBox "Trip: " & trip & " Not Found!", , "Manifest"
   frw.Visible = False: Exit Sub
 End If
DEMODE:
 DispEditMode
 frw.Visible = False
End Sub

Private Sub DispEditMode()
 Dim j As Integer, k As Integer
 Dim dm As String
 Dim b As Boolean
 
 loading = True: g2.Visible = False: g2pro.Visible = False: oldrow = -1
 Erase prolst
 ClrFrm

 trip = rs(8): assgn = rs(8): mdte = DateValue(rs(0)): odte = Format$(mdte, "DD-MMM-YYYY")
 If Len(trip) = 6 Then trpno = trip Else trpno = Mid$(trip, 2)
 
 'fix legacy trip records with 'tripnum' entry (non-prefixed numeric portion of 'assign')
 If rs(10) <> trpno Then
   dm = "update trip set tripnum='" & trpno & "' where assign='" & trip & "'"
   cmd.CommandText = dm: cmd.Execute
 End If
 
 'fix unit no. with outside owner prefix
 If rs(11) <> "" And rs(11) <> "SP" And rs(11) <> "SDT" Then 'sdt*
   t_uno = rs(11): j = Len(rs(11))
   If Left$(rs(1), j) = rs(11) Then dm = Mid$(rs(1), j + 1) Else dm = rs(1)
 Else
   dm = rs(1): t_uno = "SP" 'sdt*
 End If
 
 nottun = True: t_un = dm: nottun = False 'rs(1)
 t_un_LostFocus 'force evaluation of unit entry
 
 If t_tr = "" Then 'trailer was not automatically selected from unit (ie. straight truck)
   If Len(rs(12)) = 3 Then  'old pre-07Oct2019 3-char trailer owner codes
     If rs(12) <> "" Then
       t_tro = rs(12): j = Len(rs(12))
       If Left$(rs(7), j) = rs(12) Then dm = Mid$(rs(7), j + 1) Else dm = rs(7)
     Else
       dm = rs(7)
     End If
     trailer = dm
     j = InStr(trailer, "-") 'look for hyphen in trailer (is it a run# ?)
     If j > 0 Then
       dm = Mid$(trailer, j + 1) 'extract after hyphen
       If Val(dm) > 1 And Val(dm) < 5 Then 'looks like a run#
         trailer = Left$(trailer, j - 1) 'use only left portion for trailer
         If rs(9) > 0 Then 'run value is in trip record - will use this
           'fall-thru - run# processed below
         Else 'no run value in trip record
           trlrapp = dm 'extract run# from trailer
         End If
       End If
     End If
     If Left$(trailer, 3) = "US " Then
       trailer = Mid$(trailer, 4)
       If Len(trailer) = 6 Then ch_htr = True
     End If
   Else    'owner codes changed to SP + SCAC or equiv for others - 07Oct2019(DNY)
     If Len(Trim$(rs(12))) = 4 Then t_tro = rs(12)
     trailer = rs(7)
   End If
   notttr = True: t_tr = trailer: notttr = False
   t_tr_LostFocus 'force evaluation of trailer entry
 End If
 
 seal = rs(2): ldsht = rs(5): descr = rs(6): l_loadby = rs(17)
 orig = rs(3): dest = rs(4): driver = rs(13): rtrpn = rs(14)
 If rs(9) > 0 Then
   trlrapp = rs(9) 'always use trip record 'run' value if entered
   b = True 'indicate trip record has been created or updated by ProTrace
 Else
   b = False 'trip came from Unix sys
 End If
  '22Aug13(DNY) test for linehaul trip - set 'Arrive Linehaul' button accordingly
 Select Case Val(dest)
   Case 1, 3, 4, 5, 6, 8, 10, 11, 12, 14
     Select Case Val(orig)
       Case 1, 3, 4, 5, 6, 8, 10, 11, 12
         If orig <> dest Then
           c_arr.Visible = True: c_arr.Enabled = False
           If IsDate(Left$(rs(15), 10)) = True Then
             l_arr(0) = Format$(DateValue(Left$(rs(15), 10)), "DD-MMM-YYYY"): l_arr(1) = Mid$(rs(15), 12, 5) & "  " & rs(16)
           Else
             c_arr.Enabled = True
           End If
         End If
     End Select
 End Select
 rs.Close
 
 dm = "select pro,byndpro,ndx from trippro where tripnum='" & Right$(trip, 6) & "' order by ord"
 ''If b Then dm = dm & " order by ord" 'control order of shipments if record is updated to ProTrace version
 rsc.Open dm
 If rsc.EOF Then
   rsc.Close: loading = False: twt = "0": tsh = "0": tcu = "0": tcucnt = "": GoTo DEDIT
 End If
 
 'get shipments
 j = -2
 Do
   If j > 206 Then Exit Do
   j = j + 2: curow = j
   g2pro = rsc(0) 'trigger g2pro_Change
   k = Fix(j / 2) + 1
   prolst(k, 0) = rsc(0): prolst(k, 1) = "0" 'buff original pro#, default status is 'remove' (vs. keep/update) by default
   rsc.MoveNext
 Loop Until rsc.EOF
 rsc.Close
 g2pro.Visible = False: g2.Visible = True
 WrTotWtCube
  
DEDIT:
 'display EDIT mode
 l_ne = "EDIT": l_ne.ForeColor = &H4030B0
 c_new.Caption = "Save Changes to Trip": c_new.BackColor = &HB0BAE0    '&HD0E6E6
 c_ldtrp.Caption = "EXIT Edit": c_ldtrp.BackColor = &HB0BAE0
 c_lddte.Caption = "EXIT Edit Mode": c_lddte.BackColor = &HB0BAE0
 loading = False
 Select Case Val(tsh) 'total wt
  Case 0 '= no shipments = Trip-Only manifest
    '21Nov13 l_prnt(4) = "Trip-Only Manifest": l(23).Visible = True: t_tom.Visible = True:  frtom.Visible = True
  Case Else
    'check for Dock Pickup/Transfer manifest
    Select Case Val(orig)
      Case 37, 38, 68, 69, 71, 72, 73
       '21Nov13 21Nov13  Print Manifest button handles all actions
        c_prntdr.Visible = True '21Nov13 Print Manifest button handles all actions
      Case Else
       '21Nov13 21Nov13  Print Manifest button handles all actions
        c_prntbols.Visible = True: c_prntdr.Visible = True
        If mef Then '//^\\_//^\\
        End If
    End Select
 End Select
End Sub
Private Sub c_new_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 cnewclick = True
End Sub

Private Sub c_new_Click()
 Dim dm As String, dm2 As String, ab As String, ap As String, t As String, d As String
 Dim tripnum As String, BC As String, pndx As String
 Dim i As Integer, j As Integer, cnt As Integer
 Dim apdte As Date
 Dim b As Boolean, dlvr As Boolean, edmod As Boolean, pcxfgo As Boolean, interusa As Boolean
 Dim sp() As String
 If l_ne = "EDIT" Then edmod = True Else edmod = False
 If Len(trip) = 7 Then
   tripnum = Right$(trip, 6)
 ElseIf Len(trip) = 6 Then
   tripnum = trip
 Else
   tripnum = ""
 End If
 g2profoc = False
 'if city delivery, test for site-closed account code from weekly sched or pre-set date
 Select Case Val(orig)
   Case 1, 3, 4, 5, 6, 8, 10, 11, 12, 14
    If Val(orig) = Val(dest) Then
      For j = 0 To 206 Step 2
        If g2.TextMatrix(j, 17) <> "" And g2.TextMatrix(j, 3) <> "" Then
          If InStr(g2.TextMatrix(j + 1, 3), "SITE CLOSED") > 0 Or InStr(g2.TextMatrix(j, 6), "SITE CLSD") > 0 Then
            sqs = "  ** Consignee SITE CLOSED for Shipment " & g2.TextMatrix(j, 17) & " on " & Format$(mdte, "DDD DD-MMM-YYYY") & " **" & _
                   vbCrLf & vbCrLf & "                 Continue Save Manifest?"
            If MsgBox(sqs, vbDefaultButton2 + vbCritical + vbYesNo, "** Site Closed Alert **") = vbNo Then
              cnewclick = False: Exit Sub
            End If
          End If
        End If
      Next j
    End If
 End Select

 If savprnt Then ab = " Aborting Save/Print . ." Else ab = " Aborting Save . ."
 If orig = "" Or dest = "" Then b = True
 If Not b Then
   If lhcod(Val(orig)) = "" Then
     MsgBox "Unknown Origin Code." & ab, , "Manifest"
     If savprnt Then savprnt = False 'indicate to abort print as well
     cnewclick = False: Exit Sub
   End If
   If lhcod(Val(dest)) = "" Then
     MsgBox "Unknown Destination Code." & ab, , "Manifest"
     If savprnt Then savprnt = False
     cnewclick = False: Exit Sub
   End If
   If Trim$(t_tr) = "" Then b = True
 End If
 
 'test for invalid orig/dest combinations
 dlvr = False: interusa = False 'auto-deliver indicator, interusa manifest indicator
 If Not b Then
   'check dock pickup/transfer/storage origin codes
   
   '19May20(DNY) allow LH moves in USA between US Partner terminals
   Select Case Val(orig)
     Case 45, 46, 47, 51, 52, 53, 54, 55, 57 To 61, 63, 67  '19May2020 -> PYLE 45 to 47, EXLA 52 to 54, YRC 55, HMES 57 to 61, WARD 63, SEFL 67
       Select Case Val(dest)
         Case 45, 46, 47, 51, 52, 53, 54, 55, 57 To 61, 63, 67
           interusa = True: GoTo MAY1901 'skip to manifested shipments check
       End Select
   End Select
   
   
   Select Case Val(orig)
     Case 99 'any dest ok
     Case 31 'storage
       If orig <> dest Then
         MsgBox "Invalid Destination Code for Origin Code: Storage Trailer." & ab, , "Manifest"
         If savprnt Then savprnt = False
         cnewclick = False: Exit Sub
       End If
     Case 37, 38, 68, 69, 71, 72, 73 'all Speedy terminal dock pickup/transfer codes as origin
       If orig = dest Then
         dlvr = True
         'fall-thru where dest=orig indicates Customer pickup at Speedy terminal dock
       Else
         If Val(orig) = 68 And orig <> dest Then
           MsgBox "Destination Must Equal Origin Code for Same-Day (Code 68) Trip." & ab, , "Manifest"
           If savprnt Then savprnt = False
           cnewclick = False: Exit Sub
         End If
         Select Case Val(dest)
           Case 25 To 28, 34, 35, 45, 46, 48, 51, 64, 65, 76, 77, 78, 79, 80, 81, 82, 85, 86, 87, 88, 91, 94, 95, 96, 99 ' all agent/partner codes as destination
             'OK - fall-thru
           Case Else
             MsgBox "Invalid Destination Code for Dock Pickup/Transfer Origin Code." & ab, , "Manifest"
             If savprnt Then savprnt = False
             cnewclick = False: Exit Sub
         End Select
       End If
   End Select
   'check other destination code combinations
   Select Case Val(dest)
     Case 31
       If orig <> dest Then
         MsgBox "Invalid Origin Code for Destination Code: Storage Trailer." & ab, , "Manifest"
         If savprnt Then savprnt = False
         cnewclick = False: Exit Sub
       End If         '35-LAC as interline will become 11-LAC/VAU as SZTG terminal
     Case 25 To 28, 34, 35, 45, 46, 48, 51, 52, 53, 54, 55, 57 To 65, 76, 77, 78, 79, 80, 81, 82, 85, 86, 87, 88, 91, 94, 95, 96  ' all agent/partner codes
       Select Case Val(orig)
         Case 1, 3, 4, 5, 6, 8, 10, 12, 37, 38, 71, 72, 73, 99 'add 11-VAU once 35-LAC is retired/VAU in place  all Speedy terminal/pickup codes as origin: OK fall-thru
         Case Else
           MsgBox "Origin Code Must be a Valid Speedy Terminal Code for an Agent/Partner Destination Code." & ab, , "Manifest"
           If savprnt Then savprnt = False
           cnewclick = False: Exit Sub
       End Select
     Case 99 '10Jun20 - per PMurphy, any origin ok
   End Select
 End If
 
 If b Then
   dm = "Minimum Requirements for Manifest Not Met!" & ab & vbCrLf & vbCrLf & _
        "  (Origin, Destination, Trailer entries)"
   MsgBox dm, , "Manifest"
   If savprnt Then savprnt = False
   cnewclick = False: Exit Sub
 End If
 
 'check if trailer used more than once in day
 t = Trim$(t_tr)
 If t <> "" Then
   'fix case where full speedy trailer description is auto-found
   If InStr(t_tr, " ") > 0 Then 'if spaces in trailer entry ..
     sp = Split(t_tr, " "): t = sp(UBound(sp)) 'split by space, take last segment only
   End If
   Select Case t_tro 'owner
     Case ""    'Speedy or owner added as prefix directly to trailer#
       If IsNumeric(Left(t, 1)) = True Then 'no prefix, assume Speedy trailer
         t = CStr(Val(t)) 'remove leading zeros
       Else
         'has owner prefix - process as-is - fall-thru
       End If
     Case "SP", "SDT" 'Speedy sdt*
       If IsNumeric(t) = True Then t = CStr(Val(t)) 'remove leading zeros if non-special trailer
     Case Else 'other owner
       j = Len(t_tro)
       If Left$(t, j) = t_tro Then
         'prefix matches owner code, process trailer entry as-is
       Else 'non-match - add prefix to trailer entry
         t = t_tro & t
       End If
   End Select
   If edmod Then
     d = Format$(DateValue(mdte), "YYYY-MM-DD")
     If Val(trlrapp) = 0 Then dm = "cast(run as unsigned) < '2'" Else dm = "run='" & CStr(Val(trlrapp)) & "'"
     dm = "select cast(date_modified as char),hr_mod,min_mod,init,tripnum,origin,dest,run from trip where " & _
          "date='" & d & "' and " & dm & " and trailer='" & trailer & "' order by recno desc limit 1"
     rs.Open dm
     If Not rs.EOF Then
       If tripnum <> rs(4) Then
         dm = "! ! ! DUPLICATE TRIP ! ! ! DUPLICATE TRIP ! ! ! DUPLICATE TRIP ! ! !" & vbCrLf & vbCrLf & _
              "Trailer " & trailer & " Run " & rs(7) & " Already Used Today for Trip#: " & rs(4) & " by: " & rs(3) & vbCrLf & vbCrLf & _
              "Entered: " & Format$(DateValue(rs(0)), "DD-MMM-YYYY") & " " & Format$(rs(1), "00") & ":" & Format$(rs(2), "00") & "Hrs  " & _
              "Orig: " & rs(5) & " Dest: " & rs(6) & vbCrLf & vbCrLf & _
              "If Re-Printing an Existing Manifest, Be Aware that Any Changes to Manifest Details are Automatically Saved." & _
               vbCrLf & vbCrLf & "If Re-Assigning a Hold Trailer, Do NOT Continue - Select a Different Un-Used Trailer." & _
               vbCrLf & vbCrLf & "Do You Wish to Continue?"
         rs.Close
         If MsgBox(dm, vbDefaultButton2 + vbExclamation + vbYesNo, "Trailer Re-Assignment Check") = vbNo Then
           If savprnt Then savprnt = False
           cnewclick = False: Exit Sub
         End If
       End If
     End If
     If rs.State <> 0 Then rs.Close
   End If
   If InStr(c_new.Caption, "Changes") = 0 Then
     d = Format$(DateValue(mdte), "YYYY-MM-DD")
     '            unit seal  trip    descr  ldsht  run driver -> form fieldnames
     dm = "select unit,seal,tripnum,carrier,loader,run,driver from trip where date='" & _
           d & "' and trailer='" & t & "' and run='" & CStr(Val(trlrapp)) & _
           "' order by tripnum desc limit 1"
     rs.Open dm
     If Not rs.EOF Then
       rs.Close
       dm = "* * * * * * * * * * * * *  D U P L I C A T E  T R I P * * * * * * * * * * * *" & vbCrLf & vbCrLf & _
            "Manifest Already Exists for Trailer: " & t & ", Run: " & CStr(Val(trlrapp)) & "  on " & _
            d & vbCrLf & vbCrLf & "Please Change the *RUN No. Only* to the Next Higher value." & _
            vbCrLf & vbCrLf & "*** NOTE: Do *NOT* Change the Trailer No. (eg. do NOT add hyphen)"
       If savprnt Then savprnt = False
       MsgBox dm, , "Manifest": cnewclick = False: Exit Sub
     End If
     rs.Close
   End If
 End If
 
 '******* KIDY ***************** KINDERSLEY ********************* KIDY ********************* KINDERSLEY **********

 Select Case Val(dest)
   Case 34 'KIDY TOR terminal
     kidyfrm.Show 1
     If gen2 = "" Then
       If savprnt Then savprnt = False
       cnewclick = False: Exit Sub
     End If
 End Select
 
 '******* OMSX ***************** OMS EXPRESS ******************** OMSX ********************* OMS EXPRESS *********
 If dest = "78" Then
   Select Case Val(orig)
     Case 1, 3, 4, 5, 6, 8, 10, 12 '11-VAU when opened 'OK as linehaul from any Speedy terminal
     Case Else
       MsgBox "Invalid Origin Code for Shipment Transfer to OMS Express. Use Speedy Terminal Codes (eg. 1 - 40)", , "Manifest"
       If savprnt Then savprnt = False
       cnewclick = False: Exit Sub
   End Select
 End If
  
 '******* NCGG ***************** GARDEWINE ********************** NCGG ********************* GARDEWINE ***********
 If dest = "79" Then
   Select Case Val(orig)
     Case 1, 3, 4, 5, 6, 8, 10, 12 '11-VAU when opened 'OK as linehaul from any Speedy terminal
     Case Else
       MsgBox "Invalid Origin Code for Shipment Transfer to Gardewine NCGG. Use Speedy Terminal Codes (eg. 1 - 40)", , "Manifest"
       If savprnt Then savprnt = False
       cnewclick = False: Exit Sub
   End Select

   ncggfrm.Show 1 'pop-up NCGG pronumber (sticker pro) entry form
   If gen2 = "" Then 'user cancel
     If savprnt Then savprnt = False
     cnewclick = False: Exit Sub
   End If
 End If
  
 '******* TGBT ***************** GUILBAULT ********************** TGBT ********************* GUILBAULT ***********
 Select Case Val(dest)
   Case 77, 80, 87 'Guilbault MTL, QUE, TOR terminals
     Select Case Val(orig)
       Case 1, 3, 4, 5, 6, 8, 10, 12 '11-VAU  OK as linehaul from any Speedy terminal
       Case Else
         MsgBox "Invalid Origin Code for Shipment Transfer to GUILBAULT. Use Speedy Terminal Codes (eg. 1 to 87)", , "Manifest"
         If savprnt Then savprnt = False
         cnewclick = False: Exit Sub
     End Select
     tgbtfrm.Show 1
     If gen2 = "" Then
       If savprnt Then savprnt = False
       cnewclick = False: Exit Sub
     End If
 End Select
 
 '******* KNGSWY *************** KNGSWY ************************* KNGSWY ******************* KNGSWY **************
 If dest = "40" Then     'transfer Kingway QC - treat as linehaul
   Select Case Val(orig)
     Case 1, 3, 4, 5, 6, 8, 10, 12 '11-VAU OK as linehaul from any Speedy terminal
     Case Else
       MsgBox "Invalid Origin Code for Shipment Transfer to Kingsway Quebec. Use Speedy Terminal Codes (eg. 1 - 40)", , "Manifest"
       If savprnt Then savprnt = False
       cnewclick = False: Exit Sub
   End Select
   kngwyfrm.Show 1 'pop-up KNGWY pronumber (sticker pro) entry form
   If gen2 = "" Then 'user cancel
     If savprnt Then savprnt = False
     cnewclick = False: Exit Sub
   End If
 End If

 '******* MRTM  **************** MRTM  ************************** MRTM  ******************** MRTM  ****************
 If dest = "81" Then     'transfer Maritime Ontario - treat as linehaul
   Select Case Val(orig)
     Case 1, 3, 4, 5, 6, 8, 10, 12 '11-VAU OK as linehaul from any Speedy terminal
     Case Else
       MsgBox "Invalid Origin Code for Shipment Transfer to Maritime Ontario. Use Speedy Terminal Codes (eg. 1 - 81)", , "Manifest"
       If savprnt Then savprnt = False
       cnewclick = False: Exit Sub
   End Select
   mrtmfrm.Show 1 'pop-up MRTM pronumber (sticker pro) entry form
   If gen2 = "" Then 'user cancel
     If savprnt Then savprnt = False
     cnewclick = False: Exit Sub
   End If
 End If
 
 '******* PCXL  **************** PCXL  ************************** PCXL  ******************** PCXL  ****************
 If dest = "82" Then
   Select Case Val(orig)
     Case 1, 3, 4, 5, 6, 8, 10, 12 '11-VAU
     Case Else
       MsgBox "Invalid Origin Code for Shipment Transfer to PCXL (Winnipeg). Use Speedy Terminal Codes (eg. 1 - 81)", , "Manifest"
       If savprnt Then savprnt = False
       cnewclick = False: Exit Sub
   End Select
   'if any non-EXLA shipments, prompt for PCX sticker pro
   pcxfgo = False
   For j = 0 To 206 Step 2
     If Len(g2.TextMatrix(j, 17)) > 7 And Left$(g2.TextMatrix(j, 17), 4) <> "EXLA" Then
       pcxfgo = True: Exit For
     End If
   Next j
   If pcxfgo Then
     pcxlfrm.Show 1
     If gen2 = "" Then
       If savprnt Then savprnt = False
       cnewclick = False: Exit Sub
     End If
   End If
 End If

 '******* FASTFRATE ************* FASTFRATE ****************** FASTFRATE **************** FASTFRATE ****************
 Select Case dest
   Case "95", "96" 'transfer Fastfrate TOR, MTL - treat as linehaul
     Select Case Val(orig)
       Case 1, 3, 4, 5, 6, 8, 10, 12 '11-VAU
       Case Else
         MsgBox "Please Use Your Speedy Terminal Code as the Origin Code for Shipment Transfers to Fastfrate.", , "Manifest"
         If savprnt Then savprnt = False
         cnewclick = False: Exit Sub
     End Select
     tfastfrm.Show 1
     If gen2 = "" Then
       If savprnt Then savprnt = False
       cnewclick = False: Exit Sub
     End If
 End Select
 
 '******* WARD ****************** WARD *********************** WARD ********************* WARD *********************
 Select Case dest
   Case "63" 'southbound WARD
     Select Case Val(orig)
       Case 1, 8, 99  'OK only from Tor origin
       Case Else
         MsgBox "All WARD Southbound Originates from Code 1 TORonto Terminal.", , "Manifest"
         If savprnt Then savprnt = False
         cnewclick = False: Exit Sub
     End Select
 End Select
 
 '******* SEFL ****************** SEFL *********************** SEFL ********************* SEFL *********************
 Select Case dest
   Case "67" 'southbound SEFL
     Select Case Val(orig)
       Case 1, 3, 4, 5, 6, 8, 10, 12, 99 '11-VAU  OK from any Speedy terminal, misc.
       Case Else
         MsgBox "Please Use Your Speedy Terminal Code as Origin Code for Transfers to SEFL.", , "Manifest"
         If savprnt Then savprnt = False
         cnewclick = False: Exit Sub
     End Select
 End Select
 
 '******** EXLA * HMES * PYLE ******** ******* EXLA * HMES * PYLE ******* ********** EXLA * HMES * PYLE **********
 Select Case dest
   Case "45", "46", "52", "53", "54", "57", "58", "59" 'southbound transfer to EXLA or HMES Cleveland/Indianapolis/Toledo - treat as linehaul
     Select Case Val(orig)
       Case 1, 3, 4, 5, 6, 8, 10, 12, 99 '11-VAU  OK as linehaul from any Speedy terminal
       Case Else
         MsgBox "Please Use Your Speedy Terminal Code as the Origin Code for Shipment Transfers to EXLA, HMES or PYLE US Terminals.", , "Manifest"
         If savprnt Then savprnt = False
         cnewclick = False: Exit Sub
     End Select
 End Select
 
 '******* SZTG-RDWY************** ******* SZTG-RDWY********** *********** SZTG-RDWY**************
 Select Case dest
   Case "55" 'southbound transfer to RDWY Tonawanda NY - treat as linehaul
     Select Case Val(orig)
       Case 1, 3, 4, 5, 6, 8, 10, 12, 99 '11-VAU   OK as linehaul from any Speedy terminal
       Case Else
         MsgBox "Please Use Your Speedy Terminal Code as the Origin Code for Shipment Transfers to RDWY-Buffalo.", , "Manifest"
         If savprnt Then savprnt = False
         cnewclick = False: Exit Sub
     End Select
     yrcbffrm.Show 1 'pop-up RDWY-BUF beyond pro entry form
     If gen2 = "" Then 'user cancel
       If savprnt Then savprnt = False
       cnewclick = False: Exit Sub
     End If
 End Select
 
 '******* SZTG-EXLA************** ******* SZTG-EXLA********** *********** SZTG-EXLA**************
 Select Case dest
   Case "54" 'southbound transfer to EXLA Buffalo NY - treat as linehaul
     Select Case Val(orig)
       Case 1, 3, 4, 5, 6, 8, 10, 12, 99 '11-VAU  OK as linehaul from any Speedy terminal
       Case Else
         MsgBox "Please Use Your Speedy Terminal Code as the Origin Code for Shipment Transfers to EXLA-Buffalo.", , "Manifest"
         If savprnt Then savprnt = False
         cnewclick = False: Exit Sub
     End Select
     Select Case Val(g2.TextMatrix(0, 17))
       Case 6000000000# To 6999999999#
         exlabfrm.Show 1 'pop-up beyond pro entry form
         If gen2 = "" Then 'user cancel
           If savprnt Then savprnt = False
           cnewclick = False: Exit Sub
         End If
     End Select
 End Select
 
MAY1901:
 'check that pros are good, re-total weight & shipment count, remove/add appt date warning depending on orig/dest
 cnt = 0
 For j = 0 To 206 Step 2
   If g2.TextMatrix(j, 17) <> "" Then

     Select Case Val(g2.TextMatrix(j, 17))
       Case 6039999000# To 6039999999#
         dm = "  ***** ATTENTION ******" & vbCrLf & vbCrLf & "Pro: " & _
               g2.TextMatrix(j, 17) & " in Row: " & CStr(Fix((j + 2) / 2)) & _
              " is in * Pre-Pro * Range, Cannot Use in Manifest."
         MsgBox dm, , "Manifest"
         If savprnt Then savprnt = False
         cnewclick = False: Exit Sub
     End Select

     '15Mar10 - check if shipments exist on > 1 manifest ++++++++++++++++++++++++++++++
     'If l_ne = "NEW" And g2.TextMatrix(j, 3) <> "" Then 'entry on this line in grid
     If g2.TextMatrix(j, 3) <> "" Then
        dm = "select * from trippro, trip where trippro.pro='" & g2.TextMatrix(j, 17) & _
             "' and trippro.ocod <= '12' and trippro.dcod <= '12' and trippro.ocod=trippro.dcod" & _
             " and trip.tripnum=trippro.tripnum order by trippro.tripnum desc limit 1"
        rs.Open dm
        If Not rs.EOF Then
          If rs!tripnum <> tripnum Then
            dm = "        ***** ATTENTION ****** ATTENTION ******" & vbCrLf & vbCrLf & "Pro: " & _
                  g2.TextMatrix(j, 17) & " in Row: " & CStr(Fix((j + 2) / 2)) & _
                 " Manifested to City Trip: " & rs(1) & vbCrLf & vbCrLf & _
                 "Date: " & Format$(DateValue(rs(3)), "DD MMM YY") & "  Trailer: " & Trim$(rs(7)) & _
                 "  Orig: " & rs(10) & " " & rs(8) & "  Dest: " & rs(11) & " " & rs(9) & vbCrLf & vbCrLf & _
                 "If moving shipment to this manifest, please remove from previous manifest." & _
                  vbCrLf & "         (can be removed after saving this manifest)" & vbCrLf & _
                  vbCrLf & "              Do you wish to ABORT the SAVE?"
            If MsgBox(dm, vbDefaultButton1 + vbExclamation + vbYesNo) = vbYes Then
              If savprnt Then savprnt = False
              rs.Close: cnewclick = False: Exit Sub
            End If
          End If
        End If
        rs.Close
     End If '15-Mar-10 +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   
     If g2.TextMatrix(j, 5) <> "" Then wt = wt + Val(Replace(g2.TextMatrix(j, 5), Chr(13), ""))
     cnt = cnt + 1
     If g2.TextMatrix(j, 13) = "1" Then 'appt bill
       i = InStr(g2.TextMatrix(j, 6), Chr(13)) - 1
       If i > 6 Then
         b = False: apdte = DateValue(Left$(g2.TextMatrix(j, 6), i))
         If apdte <> mdte Then 'appt date mis-match! Only warn if local trip or drop
           Select Case Val(orig)
             Case 1 'TOR - has drops so treat special
               Select Case Val(dest)
                 Case 1, 77: b = True  'local TOR city trip or drop
               End Select
             Case 3, 4, 5, 6, 10, 8, 11, 12 'other Speedy terminals
               If Val(dest) = Val(orig) Then b = True 'only local city trips
           End Select
           Select Case InStr(g2.TextMatrix(j, 6), "Chk Date")
             Case 0 'not marked
               If b Then
                 If Not Right(g2.TextMatrix(j, 6), 1) = Chr(13) Then g2.TextMatrix(j, 6) = g2.TextMatrix(j, 6) & Chr(13)
                 g2.TextMatrix(j, 6) = g2.TextMatrix(j, 6) & "*Chk Date*"
               End If
             Case Else 'was marked
               If Not b Then g2.TextMatrix(j, 6) = Replace(g2.TextMatrix(j, 6), "*Chk Date*", "")
           End Select
         Else 'dates OK
           If InStr(g2.TextMatrix(j, 6), "Chk Date") > 0 Then g2.TextMatrix(j, 6) = Replace(g2.TextMatrix(j, 6), "*Chk Date*", "")
         End If
       End If
     End If
   End If
 Next j
 
 dockpkup = False
 If cnt = 0 Then 'no pros - just need trip# ??
   If Len(init) <> 3 Then
     MsgBox "Your Initials are from the Entry Field Top-Right.  Please Re-Enter your 3-Letter ProTrace initials to Proceed.", , "Trip Manifest"
     savprnt = False: cnewclick = False: Exit Sub
   End If
   dm = "No Shipments Specified. Create/Update a Trip-Only Manifest?"
   If MsgBox(dm, vbDefaultButton1 + vbQuestion + vbYesNo, "Save Manifest") = vbNo Then
     savprnt = False: cnewclick = False: Exit Sub
   End If

   b = False
   b = SavMnfst(2)
   
   '- --  ---  ----   ---  -- - trigger CLOSED SCHEDULE alert email if orig-dest terminals rules match
   
   '19May20(DNY) LH moves in USA between US Partner terminals
   Select Case Val(orig)
     Case 45, 46, 47, 52, 53, 54, 55, 57 To 61, 63, 67  '19May2020 -> PYLE 45 to 47, EXLA 52 to 54, YRC 55, HMES 57 to 61, WARD 63, SEFL 67
       Select Case Val(dest)
         Case 45, 46, 47, 52, 53, 54, 55, 57 To 61, 63, 67
           If b Then 'save error or user declined to save an Inter-USA manifest
             savprnt = False: cnewclick = False: Exit Sub
           End If
           'create/update closed schedule record for Inter-USA LH trip
           Select Case Val(Right$(trip, 6))
              Case 100000 To 999999 'valid trip# assigned
              Case Else
                savprnt = False: cnewclick = False: Exit Sub
           End Select
           If Len(unit) > 3 And Len(trailer) > 3 Then GoTo CTR01
       End Select
   End Select
   
   
   Select Case Val(orig)
     Case 1, 3, 4, 5, 6, 8, 10, 12, 11
       Select Case Val(dest)
         '                          'QUE & all Cdn. interline partner transfers
         Case 1, 3, 4, 5, 6, 8, 10, 12, 14, 25 To 28, 34, 35, 42, 43, 51, 76, 77, 78, 79, 80, 81, 82, 85, 86, 87, 88, 91, 94, 95, 96, 11
           If Val(orig) <> Val(dest) And (cnewclick = False Or Val(dest) = 28 Or Val(dest) = 27 Or Val(dest) = 26 Or Val(dest) = 25) Then '- - - - DEBug
CTR01:
             If t_tro = "" Then towner = "SP" Else towner = t_tro
             If rs.State <> 0 Then rs.Close
             dm = "select recno from closedtrailerreport where trip='" & Right$(trip, 6) & "' and tripdate='" & Format$(mdte, "YYYY-MM-DD") & _
                  "' and trailer='" & trailer & "' and alertsent=''"

             rs.Open dm
             If Not rs.EOF Then
               dm = "update closedtrailerreport set entrydt='" & Format$(Now, "YYYY-MM-DD Hh:Nn:Ss") & "',trip='" & Right$(trip, 6) & _
                    "',tripdate='" & Format$(mdte, "YYYY-MM-DD") & "',orig='" & orig & "',dest='" & dest & "',trailer='" & trailer & _
                    "',tractor='" & unit & "',loader='',init='" & init & "',trailerowner='" & towner & "', descrip='" & mySav(descr) & _
                    "', loadsht='" & mySav(ldsht) & "', seal='" & mySav(seal) & "', driver='" & mySav(driver) & "' where recno='" & rs!recno & "'"
             Else
               dm = "insert ignore into closedtrailerreport set entrydt='" & Format$(Now, "YYYY-MM-DD Hh:Nn:Ss") & "',trip='" & Right$(trip, 6) & "',tripdate='" & Format$(mdte, "YYYY-MM-DD") & _
                    "',orig='" & orig & "',dest='" & dest & "',trailer='" & trailer & "',tractor='" & unit & "',loader='',init='" & init & _
                     "', trailerowner='" & towner & "', descrip='" & mySav(descr) & "', loadsht='" & mySav(ldsht) & "', seal='" & mySav(seal) & "', driver='" & mySav(driver) & "'"
             End If
             rs.Close: cmd.CommandText = dm: Err = 0: cmd.Execute
             If Err <> 0 Then
               MsgBox "FAIL Write Closed Trailer Update for Trip# " & tp & " Please Save this screen and send to IT ASAP." & vbCrLf & vbCrLf & Err.Description, , "Manifest": Err = 0
             End If
           End If
       End Select
        
   Case 25 To 28, 34, 35, 77, 80, 85, 86, 87, 91
      Select Case Val(dest)
         Case 1, 3, 4, 5, 6, 8, 10, 12, 11
           If Val(orig) = 35 And Val(dest) = 11 Then
             'ignore LAC(interline) to LAC(local)
           Else
             GoTo CTR01
           End If
      End Select
   End Select

 Else
 
   Select Case Val(orig)
   
     Case 31
       b = SavMnfst(5) '0 = do not update probills or stathist, do not transfer anything to Unix
       
     Case 37, 38, 69, 71, 72, 73 'terminal dock pickup/transfer codes
       dockpkup = True
       If dlvr Then b = SavMnfst(4) Else b = SavMnfst(3) 'auto-deliver status on cust/partner pickups where orig=dest
       
       ''21Nov13 c_prnt_Click - replaced with:  21Nov13
       ''Disabl4Prn
       BC = Format$(mdte, "YYYYMMDD") & "SZTG" & Right$(trip, 6) 'create ACE style barcode for manifest & gate pass
       DrawBarCode BC, 0
       pndx = GetPrnCtrl(): l_prnt_Click 1: b = False
       arg = PrntCpy(0, pndx, Format$(Now, "DDMMMYY-HhNn"), BC, "")
       
       'print gate pass w/o re-saving (dockpkup=True)
       ''21Nov13 If Not prnttom Then c_prnttom_Click '' replaced with:  21Nov13
       PrntGatePass 2, pndx, Format$(Now, "DDMMMYY-HhNn"), BC, "0"
       
       dockpkup = False
       If dlvr Then
         MsgBox "Dock-Pickup/Transfer Manifest Saved + Shipment(s) 'Delivered' + Manifest/Gate-Pass Printed", , "Manifest"
       Else
         MsgBox "Dock-Pickup/Transfer Manifest Saved + Manifest/Gate-Pass Printed", , "Manifest"
       End If
       ''21Nov13 If Not prnttom Then 'save generated by 'Create/Edit' button
         ClrFrm        'clear manifest
         If t_un.Visible Then t_un.SetFocus 'set cursor to unit field
       ''21Nov13 End If
       cnewclick = False: Exit Sub
   
     Case 68 'SameDay trip
       If Val(dest) = 68 Then b = SavMnfst(4) 'auto-deliver
       MsgBox "SAME DAY Manifest Saved, Shipment(s) 'Delivered'", , "Manifest"
       
     Case Else
       If Not savprnt Then
         b = SavMnfst(0) '0 = do not update probills or stathist, do not transfer anything to Unix
                         '1 = combined with Print manifest + gate pass - full save & transfer
       Else
         If noprotrp Then 'trip-only manifest
           b = SavMnfst(2)
         Else
           If oprnt0 Then
             b = SavMnfst(0)
           ElseIf oprnt1 Then
             b = SavMnfst(1) 'process save/print depending on print mode
           Else
             b = SavMnfst(0)
           End If
         End If
       End If
       
   End Select
   
 End If
 
 cnewclick = False
 
 'finish up by setting the display mode
 If l_ne = "EDIT" Then
   If Not savprnt Then MsgBox "Manifest Updated Successfully.", , "Manifest"
 Else
     If ch_edsav Then 'return to edit mode after save selected
       dm = trip: ClrFrm   'buff tripnum, clear manifest
       trip = dm: c_ldtrp_Click 'regen manifest in edit mode by auto-finding by tripnum
     Else 'return to edit mode after save not selected
       ClrFrm        'clear manifest
     End If
 End If
End Sub

Private Function GetTrm$(cd$)
 If rs.State <> 0 Then rs.Close
 rs.Open "select destcity,abbr from linehaul_status_codes where code='" & cd & "'"
 If Not rs.EOF Then
   If Trim$(rs(0)) = "" Then GetTrm = rs(1) Else GetTrm = rs(0)
 End If
 rs.Close
End Function

Private Function ChkTumiWC$()
 Dim dm As String
 Dim j As Integer
  
 ChkTumiWC = ""
 For j = 0 To 206 Step 2
   If Left$(g2.TextMatrix(j, 17), 4) = "SEFL" And g2.TextMatrix(j, 3) <> "" Then 'SEFL pro entry on this line in grid . . .
     If Len(Trim$(t_un)) > 2 And Len(trailer) > 2 And Left$(g2.TextMatrix(j, 3), 5) = "TUMI " Then 'unit/trailer assigned + TUMI consignee . . .
       Select Case g2.TextMatrix(j, 23) 'test for Wstrn Cda consignee location
         Case "AB", "BC", "MB", "NU", "NT", "SK", "YT"
           If Val(dest) <> 95 Then
             dm = "    TUMI Western Canada SEFL Shipments Should Route: CFF - FASTFRATE TOR, Code 95" & vbCrLf & vbCrLf & _
                  "** NOTE: A record will be saved if Dest Code 95 is Not Chosen for this Manifest! **" & vbCrLf & vbCrLf & _
                  "    Route to Dest: " & dest & " " & l_de & " anyway ? (eg. due to special circumstances)"
             If MsgBox(dm, vbDefaultButton2 + vbCritical + vbYesNo, "TUMI Wstrn Cda SEFL Shipment Routing Check") = vbYes Then
               If MsgBox("** ARE YOU SURE ? **", vbDefaultButton2 + vbCritical + vbYesNo, "TUMI Wstrn Cda SEFL Shipment Routing Check") = vbYes Then
                 'write correspondence
                 ChkTumiWC = g2.TextMatrix(j, 17): GoTo ENDCTWC 'return with pro#
               Else
                 ChkTumiWC = "Fx95": GoTo ENDCTWC 'directive to exit manifest save
               End If
             Else
               ChkTumiWC = "Fx95": GoTo ENDCTWC
             End If
           End If
       End Select
     End If
   End If
 Next j

ENDCTWC:

End Function

Private Function ChkTrailerType(tr$) As Boolean
 Dim dm As String, dm2 As String, cttype As String
 Dim ttcnt As Integer, ttrun As Integer
 
 If rs.State <> 0 Then rs.Close
 dm = "select storage,domestic from trlrtypehistory where trailer='" & tr & "' order by recno desc limit 1"
 rs.Open dm
 If Not rs.EOF Then
 
    'check if already alerted & overridden
    If rsb.State <> 0 Then rsb.Close
    dm = "select recno from trlrtypealerts where trailer='" & tr & "' and owner='SZTG' and manifestdate='" & Format$(mdte, "YYYY-MM-DD") & _
         "' and length(reason) > '9' order by recno desc limit 1"
    rsb.Open dm
    If Not rsb.EOF Then 'already alerted/overridden with reason code
       rsb.Close: Exit Function 'return False to allow manifest save
    End If
    rsb.Close
 
    If Val(rs(0)) = 1 Then      'if marked Storage Use Only
       If Val(dest) = 31 Then   ' ok if manifested to storage code ...
          rs.Close: Exit Function
       Else
          cttype = "STORAGE"     'otherwise continue below with alert & override option
       End If
    ElseIf Val(rs(1)) = 1 Then  'if marked Domestic Use Only
       Select Case Val(dest)
          Case 1 To 51, 59, 68, 70 To 82, 86 To 88, 92 To 99: rs.Close: Exit Function   'manifested local, storage, Cdn. interline partners,  ok
          Case Else: cttype = "DOMESTIC"                                                'otherwise continue below with alert & override option
       End Select
    End If
    
    'if here then need to alert user to trailer use mis-match
    dm = "Trailer " & tr & " Has Been Marked * " & cttype & " USE ONLY *" & vbCrLf & vbCrLf & _
         "If necessary, can Override with a Reason Entry which will be Reviewed by Management." & vbCrLf & vbCrLf & _
         "                   Do you Wish to Override ?"
    If MsgBox(dm, vbCritical + vbDefaultButton2 + vbYesNo, "Storage Trailer Selected for Movement!") = vbYes Then
RECHKTType:
       dm = Trim$(InputBox("Enter Reason:", cttype & " Trailer Movement-Use Reason Statement"))
       If Len(dm) < 10 Then
          ttcnt = ttcnt + 1
          If ttcnt = 3 Then
             MsgBox "Max 3 Attempts at Reason Entry Exceeded. Aborting . . ", , "Trip Manifest"
             rs.Close: ChkTrailerType = True: Exit Function
          End If
          MsgBox "Insufficient Detail in Reason Statement, Please Re-Enter with at least 10 characters.", , cttype & " Trailer Reason Statement Entry"
          GoTo RECHKTType
       End If
       If Val(trlrapp) = 0 Then ttrun = 1 Else ttrun = Val(trlrapp)
       dm = "insert ignore into trlrtypealerts set trailer='" & tr & "',owner='SZTG',currenttype='" & cttype & "',manifestdate='" & _
             Format$(mdte, "YYYY-MM-DD") & "',run='" & CStr(ttrun) & "',alertdt='" & Format$(Now, "YYYY-MM-DD Hh:Nn:Ss") & _
            "',tripnum='" & Right$(trip, 6) & "',init='" & init & "',reason='" & myStr(dm) & "'"
       cmd.CommandText = dm: cmd.Execute 'return False to allow manifest save
    Else
       ChkTrailerType = True 'return True to prevent manifest save
    End If
    
 End If
 If rs.State <> 0 Then rs.Close
 
End Function

Private Function GetUSPartnerTermDesc$(c%)
 Dim dm As String
 Select Case c
   Case 45: dm = "PYLEBUF"
   Case 46: dm = "PYLEALB"
   Case 47: dm = "PYLEHIG"
   Case 52: dm = "EXLADET"
   Case 53: dm = "EXLACHM"
   Case 54: dm = "EXLABUF"
   Case 55: dm = "RDWYBUF"
   Case 57: dm = "HMESCLE"
   Case 58: dm = "HMESIND"
   Case 59: dm = "HMESTOL"
   Case 60: dm = "HMESBUF"
   Case 61: dm = "HMESDET"
   Case 63: dm = "WARDBUF"
   Case 67: dm = "SEFLCIN"
   Case Else: dm = "UNKNWN"
 End Select
 GetUSPartnerTermDesc = dm
End Function

Private Function SavMnfst(typ%) As Boolean 'if typ > 0 then update probill records & transfer all to Unix
 Dim sqs As String, dm As String, dm2 As String, trp As String, lhc As String, lmc As String, dm3 As String
 Dim ui As String, hr As String, min As String, ftr As String, rn As String, trlrun As String
 Dim upre As String, tpre As String, fun As String, ndxv As String, otrm As String, dtrm As String, location As String
 Dim tscac As String, oabb As String, dabb As String, townr As String 'trailer status
 Dim ofull As String, dfull As String, svccode As String
 Dim j As Integer, k As Integer, cnt As Integer
 Dim dd As Double
 Dim ndx As Variant
 Dim b As Boolean, gponly As Boolean, sameday As Boolean
 Dim correrr As Boolean, lstq As Boolean, wardsouth As Boolean, seflsouth As Boolean
     ''27Sep16 remv btlmxtor As Boolean,  npmesouth As Boolean,
 Dim nprolist(1 To 105) As String
 Dim aconn As New ADODB.Connection
 Dim acmd As New ADODB.Command
 Dim ars As New ADODB.Recordset
 Dim s() As String
 Dim lor As String, lde As String
 Dim lowesmilton As String, lmpo As String, lmsubj As String 'manifest email report body
 Dim toysrus As Integer, bbb As Integer
 Dim t As String
 Dim tumiwc As String

 Dim az As Integer, nz As Integer, azpro As String '06May2022 see near end of proc

 t = Chr(9): tumiwc = ""
 If rs.State <> 0 Then rs.Close

SKPTS2:
 Select Case t_tro
   Case "", "SDT"
     t_tro = "SP": GoTo SKPTS2
   Case "SP"
     cscac = "SZTG": cownc = "Speedy": tscac = "SZTG"
     If InStr(t_tr, " ") > 0 Then
       s = Split(t_tr, " "): trailer = s(UBound(s))
       If IsNumeric(trailer) Then trailer = Format$(Val(trailer), "0000")
     Else
       If trailer = t_tr Then
         If IsNumeric(trailer) = True Then
           trailer = Format$(Val(trailer), "0000")
         End If
       Else
         trailer = t_tr
         If IsNumeric(trailer) Then trailer = Format$(Val(trailer), "0000")
       End If
     End If
     If trailer = ltm_t And ltm_d = Format$(Now, "YYYY-MM-DD Hh:Nn") Then GoTo SKPTS '<-<-<-<-<-<-<-<-<-<-<-<-<-<-<-<-<-<-<-<-<-<-<-<-<-<-<-<-
     dm2 = vbCrLf & "By: " & init & " Trm: " & uterm & " On: " & Format$(Now, "DDD DD-MMMYY Hh:Nnam/pm") & vbCrLf & vbCrLf & _
           "Trip Date: " & Format$(mdte, "DDD DD-MMMYY") & vbCrLf
     If orig = dest Then
       If orig = "31" Or dest = "31" Then
         dm2 = dm2 & " Storage Trailer"
       ElseIf orig = "44" Or dest = "44" Then
         dm2 = dm2 & " Local Strip Trailer"
       Else
         dm2 = dm2 & "    City: " & GetTrm(orig)
       End If
     Else
       dm2 = dm2 & "Linehaul: " & GetTrm(orig) & " -> " & GetTrm(dest)
     End If
     dm2 = dm2 & vbCrLf & "Shipments: " & tsh & " TotWt: " & twt: dm3 = ""
     dm = "select svccode,trailer from trlrsvc where trailer='" & trailer & "' "
     If Len(trailer) = 4 Then dm = dm & "or trailer='ST" & trailer & "' "
     dm = dm & "and owner='SZTG'"
     rs.Open dm
     If Not rs.EOF Then
       svccode = rs(0)
       Select Case rs(0)
         Case "AN"
            ''dm = cownc & " Trailer " & rs(1) & " Dispatched w/Annual Maintenance Advisory" 'subject
            dm = cscac & " Trailer " & rs(1) & " Dispatched w/Annual Maintenance Advisory" 'subject
            dm2 = dm & dm2
            dm3 = "svctran" 'mailto-list field
         Case "RR"
            ''dm = cownc & " Trailer " & rs(1) & " Dispatched w/Repair Req'd Advisory" 'subject
            dm = cscac & " Trailer " & rs(1) & " Dispatched w/Repair Req'd Advisory" 'subject
            dm2 = dm & dm2
            dm3 = "svctrrr"
       End Select
     End If
     rs.Close
     If dm3 <> "" Then
       SendMail dm, dm2, dm3
       ltm_t = trailer: ltm_d = Format$(Now, "YYYY-MM-DD Hh:Nn")
     End If
 End Select
 
 If cscac = "SZTG" Then
   If ChkTrailerType(trailer) Then
     MsgBox "            ** WARNING **         ** WARNING **" & vbCrLf & vbCrLf & "** Manifest NOT SAVED Due to Trailer Type/Use Conflict **", , "Manifest Entry"
     GoTo ENDSAVMNFST
   End If
 End If

SKPTS:
SKPTS1:
 On Error Resume Next 'all error handling is manual until next 'On Error' statement
  
 '19May20(DNY) LH moves in USA between US Partner terminals
 Select Case Val(orig)
   Case 45, 46, 47, 52, 53, 54, 55, 57 To 61, 63, 67  '19May2020 -> PYLE 45 to 47, EXLA 52 to 54, YRC 55, HMES 57 to 61, WARD 63, SEFL 67
     Select Case Val(dest)
       Case 45, 46, 47, 52, 53, 54, 55, 57 To 61, 63, 67
         dm2 = "Create/Update Inter-USA Linehaul Schedule"
         If Val(orig) = Val(dest) Then
           dm = "        ** Origin-Destination NOT-VALID **  Aborting Save . . ." & vbCrLf & vbCrLf & _
                "Schedule is coded as a US City Delivery Run from a US Partner Terminal!"
           MsgBox dm, , dm2: SavMnfst = True: GoTo ENDSAVMNFST
         End If
         'inter USA move, force user to have complete entries for an empty, so can issue an eMail alert
         dm = vbCrLf & vbCrLf & "* Loading or Gate Pass print Not Required to Trigger Closed-Schedule Alert eMail *"
         If Len(t_un) < 4 Or Len(t_tr) < 4 Then
           dm = "Inter-USA Schedules Require Both UNIT AND TRAILER ENTRY (Alert eMail trigger)." & dm
           MsgBox dm, , dm2: SavMnfst = True: GoTo ENDSAVMNFST
         End If
         seal = GetUSPartnerTermDesc(CInt(orig)) & "-" & GetUSPartnerTermDesc(CInt(dest)) & " InterUSA"
         If Len(descr) < 3 Then
           dm = "DESCRIPTION Entry of at least 3 Chars Req'd for Inter-USA Trailer Moves."
           MsgBox dm, , dm2: SavMnfst = True: GoTo ENDSAVMNFST
         End If
         dm = "Ready to SAVE Inter-USA Linehaul Schedule . . ." & vbCrLf & vbCrLf & _
              "A Closed-Schedule Alert eMail will be Issued if New or Changed - Do you wish to Proceed?"
         If MsgBox(dm, vbDefaultButton2 + vbQuestion + vbYesNo, dm2) = vbNo Then
           SavMnfst = True: GoTo ENDSAVMNFST 'return True to 'c_New' proc indicating save declined
         End If
         GoTo STSAVM01  'all-good - create/edit schedule
     End Select
 End Select
    
 ' Tumi Western CDA routing, force CFF
 tumiwc = ChkTumiWC()
 If tumiwc = "Fx95" Then GoTo ENDSAVMNFST 'user chose to abort and correct the dest code to 95
 'see further down below for dest code override + correspondence write
  
STSAVM01:
  
 'if creating, get next trip no. by running mySQL *thread-safe* next-number operation on single record table 'tripcntr'
 If l_ne = "NEW" Then
   trp = GetNxtTrpNum()
   If trp = "" Then GoTo ENDSAVMNFST
 End If
  
 'check orig/dest codes, get trip prefix code for trip
 lhc = GetTripPrefix
 
 'get 'M'anifested vs. 'L'inehaul status from orig/dest codes
 lmc = SetLMStatus
   
 'full trip code = prefix + trip next-number
 If l_ne = "NEW" Then trpno = trp
 trip = lhc & trpno 'write trip# and full prefixed assign#
   
 'save header to table 'trip'
 If l_ne = "NEW" Then
   sqs = "insert ignore into trip set tripnum='" & trpno & "',"
 Else
   sqs = "update trip set"
 End If
 ui = Left$(init, 1) & Right$(init, 1)
 hr = Format$(Now, "Hh"): min = Format$(Now, "Nn") 'hour/min last-modified
 
 wardsouth = False: seflsouth = False  ''27Sep16 npmesouth = False: btlmxtor = False
 
 arg = FixUnitTrailerRun(upre, fun, tpre, ftr, rn, trlrun) 'vars set by function
    
 '19May20(DNY) LH moves in USA between US Partner terminals
 Select Case Val(orig)
   Case 45, 46, 47, 52, 53, 54, 55, 57 To 61, 63, 67  '19May2020 -> PYLE 45 to 47, EXLA 52 to 54, YRC 55, HMES 57 to 61, WARD 63, SEFL 67
     Select Case Val(dest)
       Case 45, 46, 47, 52, 53, 54, 55, 57 To 61, 63, 67: GoTo STSAVM02
     End Select
 End Select
    
 Select Case Val(orig)
 
   Case 1
     If Val(dest) = 63 Then wardsouth = True
     If Val(dest) = 67 Then seflsouth = True

   Case 37
     If Trim$(seal) = "" Then seal = "PKUP WIN DOCK"
     If Trim$(descr) = "" Then descr = "WINDSOR DOCK PKUP"
   Case 38
     If Trim$(seal) = "" Then seal = "PKUP BRO DOCK"
     If Trim$(descr) = "" Then descr = "BROCKVILLE DOCK PKUP"  '909#

   Case 69
     If typ = 4 Then
       If Trim$(seal) = "" Then seal = "RDWY PKUP TOR DOCK"
       If Trim$(descr) = "" Then descr = "RDWY BRAMPTON DOCK PKUP"
     End If
   Case 71
     If Trim$(seal) = "" Then seal = "PKUP TOR DOCK"
     If Trim$(descr) = "" Then descr = "BRAMPTON DOCK PKUP"
   Case 72
     If Trim$(seal) = "" Then seal = "PKUP MTL DOCK"
     If Trim$(descr) = "" Then descr = "MONTREAL DOCK PKUP"
   Case 73
     If Trim$(seal) = "" Then seal = "PKUP LON DOCK"
     If Trim$(descr) = "" Then descr = "LONDON DOCK PKUP"
   Case 74
     If Trim$(seal) = "" Then seal = "PKUP PIC DOCK"
     If Trim$(descr) = "" Then descr = "PICKERING DOCK PKUP"
     
 End Select
 
 Select Case Val(dest)
   Case 67: If Val(orig) < 19 Then seflsouth = True
 End Select
 
STSAVM02:
 
 '%%%%%%%%%%%%%%%%%%% SAVE TRIP HEADER %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
 sqs = sqs & " assign='" & trip & "', date='" & Format$(mdte, "YYYY-MM-DD") & "', unit='" & fun & _
             "', trailer='" & ftr & "', run='" & rn & "', seal='" & seal & "', dest_city='" & _
              mySav(destcty(Val(dest))) & "', origin_city='" & mySav(destcty(Val(orig))) & "', distcode='" & _
              ui & "', loader='" & mySav(ldsht) & "', carrier='" & mySav(descr) & "', totalwt='" & _
              twt & "', origin='" & orig & "', dest='" & dest & "', date_modified='" & _
              Format$(Now, "YYYY-MM-DD") & "', hr_mod='" & hr & "', min_mod='" & min & "', init='" & ui & _
              "', unitpre='" & upre & "', trlrpre='" & tpre & "', driver='" & Trim$(driver) & "'"
 If l_ne = "EDIT" Then sqs = sqs & " where tripnum='" & trpno & "'"
 cmd.CommandText = sqs: Err = 0: cmd.Execute   '//^\\
 If Err <> 0 Then
   MsgBox "FAIL Save Trip Header!! Aborting Save . ." & vbCrLf & vbCrLf & Err.Description
   Err = 0: SavMnfst = True: GoTo ENDSAVMNFST 'indicate Fn returned with an error (or user declined to save manifest)
 End If
 '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
 '** Do not Transfer Trip Header to Unix Until all Pro transfers are completed OK **
  
  '19May20(DNY) LH moves in USA between US Partner terminals
 Select Case Val(orig)
   Case 45, 46, 47, 52, 53, 54, 55, 57 To 61, 63, 67  '19May2020 -> PYLE 45 to 47, EXLA 52 to 54, YRC 55, HMES 57 to 61, WARD 63, SEFL 67
     Select Case Val(dest)
       Case 45, 46, 47, 52, 53, 54, 55, 57 To 61, 63, 67
         SavMnfst = False: GoTo ENDSAVMNFST 'schedule saved no errors, done so jump out
     End Select
 End Select
   
 'setup correspondence write for SAMEDAY shipment(s), removed shipments (after a full save)
 Err = 0: aconn.Open aconnstr
 If Err <> 0 Then
   MsgBox "FAIL Open Correspondence. Save/Print Will Continue." & vbCrLf & vbCrLf & Err.Description, , "Manifest"
   Err = 0: sameday = False: correrr = True
 Else
   If Val(orig) = 68 And Val(dest) = 68 Then sameday = True
 End If
 ars.ActiveConnection = aconn: acmd.ActiveConnection = aconn: acmd.CommandType = adCmdText
 
 If tumiwc <> "" Then 'overriding CFF, write a correspondence vs. pro# returned
   acmd.CommandText = "update ndx_counter set ndx=last_insert_id(ndx+1)"
   Err = 0: acmd.Execute: ars.Open "select last_insert_id()"
   If Err = 0 Then ndxv = ars(0)
   If ars.State <> 0 Then ars.Close
   If Err <> 0 Then
     MsgBox "FAIL Get Attach Index!" & vbCrLf & Err.Description, , "Manifest": GoTo SKIPRMVPRO
   End If
   dm = "Routing to CFF Overrided in Manifest Trip# " & trpno & " for SEFL TUMI Wstrn Cda Consignee by " & init
   sqs = "insert into filemaster set ndx='" & ndxv & "', pro='" & tumiwc & _
         "', filetype='Comment', adddate='" & Format$(Now, "YYYY-MM-DD") & _
         "', addtime='" & Format$(Now, "Hh:Nn:Ss") & "', comment='" & mySav(dm) & _
         "', category='Invoice/Collection', init='ZZZ', lastevent='AddCom'"
   acmd.CommandText = sqs: Err = 0: acmd.Execute
 End If
 
 'save shipments to 'trippro' and update probill records
 cnt = 0: Erase nprolist: lowesmilton = "": toysrus = 0: bbb = 0
 
 For j = 0 To 206 Step 2
 
   If g2.TextMatrix(j, 17) <> "" And g2.TextMatrix(j, 3) <> "" Then 'entry on this line in grid

     '1. create/update 'trippro' records in ProTrace
     If rs.State <> 0 Then rs.Close
     rs.Open "select ndx from trippro where pro='" & g2.TextMatrix(j, 17) & "' and tripnum='" & trpno & "'"
     If Not rs.EOF Then 'record exists for this pro-trip
       b = True: ndx = rs(0) 'indicate edit mode, ndx is non-zero if originated/last_modified in Unix system
       sqs = "update trippro set" 'start edit SQL command string
     Else 'no existing record, start building SQL insert record command
       b = False
       sqs = "insert ignore into trippro set ndx='0', pro='" & g2.TextMatrix(j, 17) & "', tripnum='" & trpno & "',"
     End If
     rs.Close
     cnt = cnt + 1
     'continue SQL command string build
     '909#
     Select Case l_or
       Case "TOR": lor = "TR"
       Case "MTL": lor = "MT"
       Case "LON": lor = "LN"
       Case "BRO": lor = "AX"
       Case "PIC": lor = "PI"
       Case "BRA": lor = "BR"
       Case "WIN": lor = "WR"
       Case "LAC", "VAU": lor = "VA"
       Case "QCY": lor = "QC"
       Case "QUE": lor = "QC"
       Case "MIS": lor = "MI"
       Case "MIL": lor = "ML"
       Case Else: lor = l_or
     End Select
       
     If Len(trailer) > 2 And Val(dest) = 1 Then
        dm = "|" & UCase$(Replace(g2.TextMatrix(j, 3), Chr(13), "|"))
     
        'count BBB shipments
     
        '1. count up # of ToysRUs Concord inbound DC - send alert mail to DriverCheck-In team to ID the manifests to use for delivering-off
        If savprnt And Left$(dm, 4) = "|TOY" And InStr(dm, "CONCORD  ON  ") > 0 Then toysrus = toysrus + 1
     
        '2. write Lowes delivery manifests to staging table for separate central program to build/email spreadsheet manifests, test tripnum = 376621
        If savprnt And InStr(dm, "8450 BOSTON CHURCH") > 0 And InStr(dm, "MILTON ") > 0 And InStr(dm, " ON ") > 0 Then
           If rsc.State <> 0 Then rsc.Close
           rsc.Open "select po from pro_po where pro='" & g2.TextMatrix(j, 17) & "' limit 1"
           If rsc.EOF Then lmpo = "N/A" Else lmpo = Trim$(rsc(0))
           rsc.Close
           dm = "select sendcount from emanifest where pro='" & g2.TextMatrix(j, 17) & "'"
           If rsc.State <> 0 Then rsc.Close
           rsc.Open dm
           If rsc.EOF Then
             dm = "insert ignore into emanifest set partnerid='LOWES',tripnum='" & trpno & "',tripdate='" & Format$(mdte, "YYYY-MM-DD") & _
                  "',trailer='" & trailer & "',triprun='" & rn & "',unit='" & fun & "',pro='" & g2.TextMatrix(j, 17) & "',po='" & lmpo & _
                  "',pcs='" & Trim$(Replace(g2.TextMatrix(j, 4), Chr(13), "")) & "',wt='" & Trim$(Replace(g2.TextMatrix(j, 5), Chr(13), "")) & _
                  "',shipper='" & Trim$(Replace(g2.TextMatrix(j, 2), Chr(13), "")) & "',sendtrigger='G',manifestinits='" & init & "'"
           Else
             dm = "update emanifest set partnerid='LOWES',tripnum='" & trpno & "',tripdate='" & Format$(mdte, "YYYY-MM-DD") & _
                  "',trailer='" & trailer & "',triprun='" & rn & "',unit='" & fun & "',po='" & lmpo & _
                  "',pcs='" & Trim$(Replace(g2.TextMatrix(j, 4), Chr(13), "")) & "',wt='" & Trim$(Replace(g2.TextMatrix(j, 5), Chr(13), "")) & _
                  "',shipper='" & Trim$(Replace(g2.TextMatrix(j, 2), Chr(13), "")) & "',sendtrigger='G',manifestinits='" & init & _
                  "' where pro='" & g2.TextMatrix(j, 17) & "'"
           End If
           'Clipboard.Clear: Clipboard.SetText dm
           cmd.CommandText = dm: Err = 0: cmd.Execute
           If Err <> 0 Then
             dm = "** FAIL Set LOWES Auto-eManifest eMail Trigger ** Lowes SOP Violation! Please Advise IT Immediately!"
             MsgBox dm, , "Trip Manifest": Err = 0
           End If
        End If
     End If
     '------------------------------------ End Lowes Milton --------------------------------------------------------------------------------------------------------
     
     
     sqs = sqs & " trip='" & trip & "', date='" & Format$(Now, "YYYY-MM-DD") & "', hr='" & hr & _
                 "', min='" & min & "', unit='" & fun & "', trailer='" & ftr & "', orig='" & _
                  lor & "', dest='" & l_de & "', ocod='" & orig & "', dcod='" & dest & _
                 "', run='" & rn & "', ord='" & CStr(cnt) & "', tripdate='" & Format$(mdte, "YYYY-MM-DD") & _
                 "', hazmat='" & g2.TextMatrix(j, 16) & "', byndpro='" & g2.TextMatrix(j, 20) & "'"
     'complete SQL command string for edit mode
     If b Then sqs = sqs & " where pro='" & g2.TextMatrix(j, 17) & "' and tripnum='" & trpno & "'"
     cmd.CommandText = sqs: Err = 0: cmd.Execute  '//^\\
     If Err <> 0 Then
       dm = "FAIL Save Pro " & g2.TextMatrix(j, 17) & " at Row " & CStr(Fix(j / 2) + 1) & " to 'trippro'" & _
             vbCrLf & vbCrLf & " ** Save Aborted **" & vbCrLf & vbCrLf & Err.Description
       Err = 0: MsgBox dm, , "Manifest": GoTo ENDSAVMNFST
     End If
     
     If testo = "" Then testo = g2.TextMatrix(j, 17)
  
     'Kingsway, EXLA-Buffalo, YRCRoadway, Guilbault, Fastfrate, Meyers*89*gone, NewPenn beyond pro , add Ward 07Apr15, add YRCRoadway 18Mar16
     Select Case Val(dest)
       Case 34, 40, 55, 63, 77, 79, 80, 87, 95, 96  '09May17 EXLA-SZTG effectively the same, 23Oct18-Gardewine/NCGG/79 - beyond pro save
         If Len(Trim$(g2.TextMatrix(j, 20))) > 5 Then
           sqs = "update probill set agentpro='" & g2.TextMatrix(j, 20) & "' where pronumber='" & g2.TextMatrix(j, 17) & "'"
           cmd.CommandText = sqs: Err = 0: cmd.Execute: Err = 0
         End If
     End Select
     
     '2. if editing and this pro was in original manifest, then mark original pro list to 'keep' pro
     If b Then 'if editing . .
       For k = 1 To 104 'cycle thru trip's original pro list
         If prolst(k, 0) <> "" And g2.TextMatrix(j, 17) = prolst(k, 0) Then 'match found
           prolst(k, 1) = "1": GoTo FPLPRO 'change status from 'remove' to 'keep'
         End If
       Next k
       'if here then pro was not in original list, was added to manifest this session
     End If
     nprolist(cnt) = g2.TextMatrix(j, 17) 'buff for all new/updated/removed pros for Unix-transfer processing

FPLPRO:
     '************** '//^\\_//^\\
     ''21Nov13  If typ > 0 Or (mef And typ = 0 And o_prnt(2).Value) Then 'only update pros & transfer to Unix if generating full set and/or gate pass
     If typ > 0 Then  ''Or (mef And typ = 0 And oprnt2) Then 'only update pros & transfer to Unix if generating full set and/or gate pass
     '**************

       If sameday Then 'check if SameDay already saved to prevent a save for multiple clicks on 'NEW'
         ars.Open "select ndx from filemaster where pro='" & g2.TextMatrix(j, 17) & _
                  "' and adddate='" & Format$(Now, "YYYY-MM-DD") & _
                "' and comment='SAME DAY Shipment. Invoice Accordingly.'"
         If Not ars.EOF Then
           ars.Close: GoTo NXTSAVJ
         End If
         ars.Close
       End If

       '3. update probill records in both ProTrace and Unix
       '   *** field 'edi_trip' maps to Unix table 'probilln' field 'trip'
       '       populating this field allows editing of normally protected fields
       '   set 'cc_broker' to blank to signal 'prosqlget' Unix program to process 'delivery', 'loadno' fields
       Select Case Val(dest)
         Case 31
           ''sqs = "update probill set delivery='" & trlrun & "', funds='S' where pronumber='" & g2.TextMatrix(j, 17) & "'" 'deprecated 07Oct2019(DNY)
           sqs = "update probill set delivery='" & ftr & "', funds='S' where pronumber='" & g2.TextMatrix(j, 17) & "'"      'write owner-SCAC + trailer# to Probill View
         Case 34
           arg = arg
           
         Case 40 'KNGSWY linehaul    trlrun replaced with ftr 07Oct19
           sqs = "update probill set delivery='" & ftr & "', loadno='" & CStr(cnt) & "', funds='" & lmc & _
                 "', edi_trip='" & trpno & ", dlvytrip='" & trip & "', cc_broker='MNFST', agentpro='" & _
                  g2.TextMatrix(j, 20) & "', hazmat_l3='" & mySav(g2.TextMatrix(j, 21)) & _
                 "' where pronumber='" & g2.TextMatrix(j, 17) & "'"
         Case 44 'strip - fall thru
           sqs = "NOSAVE"
         Case 77, 80, 87 'TGBT transfer - set 204 send trigger
           If Len(Trim$(g2.TextMatrix(j, 20))) > 5 Then
             sqs = "update trippro set 204sent='1', ndx='" & g2.TextMatrix(j, 20) & "' where pro='" & g2.TextMatrix(j, 17) & "'"
             cmd.CommandText = sqs: Err = 0: cmd.Execute
           End If
           If Err <> 0 Then
             dm = "FAIL Set Guilbault EDI 204 Send Trigger! ** Please Inform IT ASAP **" & vbCrLf & Err.Description
             Err = 0: MsgBox dm, , "Manifest": GoTo ENDSAVMNFST
           End If
           sqs = "update probill set deldate='" & Format$(mdte, "YYYY-MM-DD") & "', delivery='" & _
           ftr & "', loadno='" & Trim$(dest) & "', funds='" & lmc & "', edi_trip='" & trpno & _
           "', dlvytrip='" & trip & "', cc_broker='MNFST"
           If Len(Trim$(g2.TextMatrix(j, 20))) > 5 Then
             sqs = sqs & "', agentpro='" & g2.TextMatrix(j, 20)
           End If
           sqs = sqs & "', hazmat_l3='" & mySav(g2.TextMatrix(j, 21)) & "' where pronumber='" & g2.TextMatrix(j, 17) & "'"
                      
         Case 78, 79 'OMS Express, Gardewine - replaced MDS N.Ont. coverage starting 22Oct2018
           sqs = "update probill set deldate='" & Format$(mdte, "YYYY-MM-DD") & "', delivery='" & _
           ftr & "', loadno='" & Trim$(dest) & "', funds='" & lmc & "', edi_trip='" & trpno & _
           "', dlvytrip='" & trip & "', cc_broker='MNFST"
           If Len(Trim$(g2.TextMatrix(j, 20))) > 5 Then
             sqs = sqs & "', agentpro='" & g2.TextMatrix(j, 20)
           End If
           sqs = sqs & "', hazmat_l3='" & mySav(g2.TextMatrix(j, 21)) & "' where pronumber='" & g2.TextMatrix(j, 17) & "'"
         
         Case 81 'MRTM linehaul - set trigger for 204 send, write beyond pro + populate loadno field with dest code for legacy
           sqs = "update trippro set 204sent='1' where pro='" & g2.TextMatrix(j, 17) & "'"
           cmd.CommandText = sqs: Err = 0: cmd.Execute
           If Err <> 0 Then
             dm = "FAIL Set EDI 204 Send Trigger! ** Please Inform IT ASAP **" & vbCrLf & Err.Description
             Err = 0: MsgBox dm, , "Manifest": GoTo ENDSAVMNFST
           End If
           sqs = "update probill set deldate='" & Format$(mdte, "YYYY-MM-DD") & "', delivery='" & _
           ftr & "', loadno='" & Trim$(dest) & "', funds='" & lmc & "', edi_trip='" & trpno & _
           "', dlvytrip='" & trip & "', cc_broker='MNFST', agentpro='" & g2.TextMatrix(j, 20) & _
           "', hazmat_l3='" & mySav(g2.TextMatrix(j, 21)) & "' where pronumber='" & g2.TextMatrix(j, 17) & "'"
         Case 95, 96 'Fastfrate TOR, MTL - not sending 204 as of 07Jul14
           sqs = "update probill set deldate='" & Format$(mdte, "YYYY-MM-DD") & "', delivery='" & ftr & "', loadno='" & Trim$(dest) & _
                 "', funds='" & lmc & "', edi_trip='" & trpno & "', dlvytrip='" & trip & "', cc_broker='MNFST"
           If Len(Trim$(g2.TextMatrix(j, 20))) > 5 Then sqs = sqs & "', agentpro='" & g2.TextMatrix(j, 20)
           sqs = sqs & "', hazmat_l3='" & mySav(g2.TextMatrix(j, 21)) & "' where pronumber='" & g2.TextMatrix(j, 17) & "'"
           
         Case Else
           If typ = 4 Then 'both update probill And auto-deliver - driven only by origin codes
             sqs = "update probill set deldate='" & Format$(mdte, "YYYY-MM-DD") & "', delivery='" & _
                    ftr & "', loadno='" & CStr(cnt) & "', funds='D', edi_trip='" & trpno & _
                    "', dlvytrip='" & trip & "', cc_broker='MNFST' where pronumber='" & g2.TextMatrix(j, 17) & "'"
           Else
             If wardsouth Then '1 to 63 trip
               sqs = "update probill set deldate='" & Format$(mdte, "YYYY-MM-DD") & "', delivery='" & _
                      ftr & "', loadno='" & CStr(cnt) & "', funds='" & lmc & "', edi_trip='" & trpno & _
                      "', dlvytrip='" & trip & "', cc_broker='MNFST' where pronumber='" & g2.TextMatrix(j, 17) & "'"
             Else 'ALL OTHER TRIPS ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
               sqs = "update probill set deldate='" & Format$(mdte, "YYYY-MM-DD") & "', delivery='" & _
                      ftr & "', loadno='" & CStr(cnt) & "', funds='" & lmc & "', edi_trip='" & trpno & _
                      "', dlvytrip='" & trip & "', cc_broker='MNFST' where pronumber='" & g2.TextMatrix(j, 17) & "'"
             End If
           End If
       End Select
       If sqs <> "NOSAVE" Then '^^
         cmd.CommandText = sqs: Err = 0: cmd.Execute  '//^\\
       Else
         typ = 6: Err = 0
       End If
       If Err <> 0 Then
         dm = "FAIL Update Probill for Pro " & g2.TextMatrix(j, 17) & " at Row " & CStr(Fix(j / 2) + 1) & _
               vbCrLf & vbCrLf & " ** Save Aborted ** Please Contact IT ASAP." & vbCrLf & vbCrLf & _
               Err.Description
         Err = 0: MsgBox dm, , "Manifest": GoTo ENDSAVMNFST
       Else
'//\\DEBUG
         If typ < 5 Then 'Not storage trailer, not local strip
         
           '=v== this section below should not run since typ must be > 0 in this block == code here is a catchall ===================================
           If Not savprnt Then '21Nov13 save initiated from Create/Edit Manifest button, not Print Manifest
             sqs = "select recno,trailer,statuscode from stathist where pro='" & g2.TextMatrix(j, 17) & "' and date_status='" & _
                     Format$(mdte, "YYYY-MM-DD") & "' and statuscode in ('M','Q','O') and origin='" & orig & "' and dest='" & dest & _
                    "' and trailer='" & ftr & "' order by date_ent desc, cast(hr_ent as unsigned) desc, cast(min_ent as unsigned) desc limit 1"
             rs.Open sqs
             If rs.EOF Then 'no prior Manifest or Loading record - create 'M'anifested status for internal record - not used for EDI, web
                sqs = "insert ignore into stathist set pro='" & g2.TextMatrix(j, 17) & "', hr_ent='" & _
                       Format$(Now, "Hh") & "', min_ent='" & Format$(Now, "Nn") & "', statuscode='M', date_ent='" & _
                       Format$(Now, "YYYY-MM-DD") & "', date_status='" & Format(mdte, "YYYY-MM-DD") & _
                      "', trailer='" & ftr & "', origin='" & orig & "', dest='" & dest & "'"
             Else
                sqs = "update stathist set hr_ent='" & Format$(Now, "Hh") & "', min_ent='" & Format$(Now, "Nn") & "', statuscode='" & rs(1) & _
                      "', date_ent='" & Format$(Now, "YYYY-MM-DD") & "', date_status='" & Format(mdte, "YYYY-MM-DD") & "', trailer='" & ftr & _
                      "', origin='" & orig & "', dest='" & dest & "' where recno='" & rs(0) & "'"
             End If
             rs.Close: cmd.CommandText = sqs: Err = 0: cmd.Execute
           End If
           '=^== this section above should not run since typ must be > 0 in this block == code here is a catchall ===================================
           
           If typ > 0 And savprnt Then
           End If
         End If
       End If
       
       If typ = 5 Then 'check if pro already on trailer on trip date
         sqs = "select recno from stathist where pro='" & g2.TextMatrix(j, 17) & "' and statuscode='S' and " & _
               "date_status = '" & Format$(mdte, "YYYY-MM-DD") & "' and trailer='" & ftr & "' and " & _
               "origin='31' and dest='31' order by recno desc limit 1"
         rs.Open sqs
         If Not rs.EOF Then
           rs.Close: GoTo SKIP68A
         End If
         rs.Close
         sqs = "insert ignore into stathist set pro='" & g2.TextMatrix(j, 17) & "', hr_ent='" & _
                Format$(Now, "Hh") & "', min_ent='" & Format$(Now, "Nn") & "', statuscode='S', date_ent='" & _
                Format$(Now, "YYYY-MM-DD") & "', date_status='" & Format(mdte, "YYYY-MM-DD") & _
               "', trailer='" & ftr & "', origin='" & orig & "', dest='" & dest & "'"
               cmd.CommandText = sqs: Err = 0: cmd.Execute '21Nov13
       End If
       If typ = 6 Then '^^
         
       End If
       If Err <> 0 Then
         dm = "FAIL Write 'M'anifested Status, Pro " & g2.TextMatrix(j, 17) & _
              " . Save/Print Will Continue." & vbCrLf & vbCrLf & Err.Description
         Err = 0: MsgBox dm, , "Manifest"
       Else
         If typ = 4 Then 'auto-deliver specified (eg. Cust Pickup at Dock manifest)
           sqs = "insert ignore into stathist set pro='" & g2.TextMatrix(j, 17) & "', hr_ent='" & _
                  Format$(Now, "Hh") & "', min_ent='" & Format$(Now, "Nn") & "', statuscode='D', date_ent='" & _
                  Format$(Now, "YYYY-MM-DD") & "', date_status='" & Format(mdte, "YYYY-MM-DD") & _
                 "', trailer='" & ftr & "', origin='" & orig & "', dest='" & dest & "', statusseq='22'"
           cmd.CommandText = sqs: Err = 0: cmd.Execute
           If Err <> 0 Then
             dm = "FAIL Write 'D'elivered Status, Pro " & g2.TextMatrix(j, 17) & _
                  " . Save/Print Will Continue." & vbCrLf & vbCrLf & Err.Description
             Err = 0: MsgBox dm, , "Manifest"
           End If
           
           If sameday Then 'write correspondence for SAMEDAY shipment(s)
             acmd.CommandText = "update ndx_counter set ndx=last_insert_id(ndx+1)"
             Err = 0: acmd.Execute: ars.Open "select last_insert_id()"
             If Err = 0 Then ndxv = ars(0)
             If ars.State <> 0 Then ars.Close
             If Err <> 0 Then
               MsgBox "FAIL Get Attach Index!" & vbCrLf & Err.Description, , "Manifest": GoTo SKIP68A
             End If
             dm = "SAME DAY Shipment. Invoice Accordingly."
             sqs = "insert into filemaster set ndx='" & ndxv & "', pro='" & g2.TextMatrix(j, 17) & _
                   "', filetype='Comment', adddate='" & Format$(Now, "YYYY-MM-DD") & _
                   "', addtime='" & Format$(Now, "Hh:Nn:Ss") & "', comment='" & dm & _
                   "', category='Invoice/Collection', init='" & init & "', lastevent='AddCom'"
             acmd.CommandText = sqs: Err = 0: acmd.Execute
             If Err <> 0 Then
               MsgBox "FAIL Save Correspondence!" & vbCrLf & Err.Description, , "Manifest": Err = 0
             End If
           End If
           
SKIP68A:
         End If
       End If
     End If
     
     If typ = 0 And (Not savprnt Or oprntL) Then '21Nov13 save initiated from Create/Edit Manifest button, not Print Manifest
         
        sqs = "select recno,trailer,statuscode from stathist where pro='" & g2.TextMatrix(j, 17) & "' and date_status='" & _
                Format$(mdte, "YYYY-MM-DD") & "' and statuscode in ('M','Q','O') and origin='" & orig & "' and dest='" & dest & _
               "' and trailer='" & ftr & "' order by date_ent desc, cast(hr_ent as unsigned) desc, cast(min_ent as unsigned) desc limit 1"
        rs.Open sqs
        
        If rs.EOF Then 'no prior Manifest or Loading record - create 'M'anifested status for internal record - not used for EDI, web
           sqs = "insert ignore into stathist set pro='" & g2.TextMatrix(j, 17) & "', hr_ent='" & _
                  Format$(Now, "Hh") & "', min_ent='" & Format$(Now, "Nn") & "', statuscode='M', date_ent='" & _
                  Format$(Now, "YYYY-MM-DD") & "', date_status='" & Format(mdte, "YYYY-MM-DD") & _
                 "', trailer='" & ftr & "', origin='" & orig & "', dest='" & dest & "'"
        Else
           sqs = "update stathist set hr_ent='" & Format$(Now, "Hh") & "', min_ent='" & Format$(Now, "Nn") & "', statuscode='" & rs(2) & _
                 "', date_ent='" & Format$(Now, "YYYY-MM-DD") & "', date_status='" & Format(mdte, "YYYY-MM-DD") & "', trailer='" & ftr & _
                 "', origin='" & orig & "', dest='" & dest & "' where recno='" & rs(0) & "'"
        End If
        rs.Close: cmd.CommandText = sqs: Err = 0: cmd.Execute
        
        sqs = "update probill set deldate='" & Format$(mdte, "YYYY-MM-DD") & "', delivery='" & _
               ftr & "', loadno='" & CStr(cnt) & "', funds='" & lmc & "', edi_trip='" & trpno & _
              "', dlvytrip='" & trip & "', cc_broker='MNFST' where pronumber='" & g2.TextMatrix(j, 17) & "'"
              cmd.CommandText = sqs: Err = 0: cmd.Execute
     End If
   End If
NXTSAVJ:
 Next j
 
 If toysrus > 0 Then
   dm = "ToysRUs DC Manifest Trip# " & trpno & " - " & CStr(toysrus) & " Shipments": SendMail dm, vbCrLf & dm, "ToysRUsManifestAlert"
 End If
 
 '06May2022(DNY) detect full-loads consigned to an AMAZON location on city trip '' test trip#  234331
 Select Case Val(orig)
    Case 1, 3, 4, 5, 6, 8, 10, 11, 12 'city trip
       If Val(orig) = Val(dest) Then
          az = 0: nz = 0
          For j = 0 To 206 Step 2
             Select Case Left$(UCase$(g2.TextMatrix(j, 3)), 6) 'col 3 = consignee name/address/city-prov-postzip
               Case ""
               Case "AMAZON"
                 dm = Right$(UCase$(Trim$(g2.TextMatrix(j, 3))), 6)
                 If az = 0 Then
                    azpc = dm: az = 1: azpro = g2.TextMatrix(0, 17)
                 Else
                    If dm = azpc Then az = az + 1 Else nz = nz + 1
                 End If
               Case Else: nz = nz + 1
             End Select
          Next j
          If az > 0 And nz = 0 Then
             If rs.State <> 0 Then rs.Close
             dm = "select recno,active from closedfullloadschedules where tripnum='" & Right$(trip, 6) & "' and amazon='1'"
             rs.Open dm
             If rs.EOF Then
                dm = "insert ignore into closedfullloadschedules set tripnum='" & Right$(trip, 6) & "', amazon=1, " & _
                     "trailer='" & trailer & "', active='1', dlvyterm='" & l_or & "', totwt='" & twt & "', shipcnt='" & tsh & _
                     "', inits='" & init & "', lastmodified='" & Format$(Now, "YYYY-MM-DD Hh:Nn:Ss") & "', firstpro='" & azpro & "'"
                cmd.CommandText = dm: Err = 0: cmd.Execute
             Else
                If Val(rs!active) = 9 Then 'check if schedule was delivered/closed-out
                   'skip - no changes to exiting record
                Else
                  dm = "update closedfullloadschedules set trailer='" & trailer & "', dlvyterm='" & l_or & "', totwt='" & twt & _
                        "', shipcnt='" & tsh & "', inits='" & init & "',  lastmodified='" & Format$(Now, "YYYY-MM-DD Hh:Nn:Ss") & _
                        "', apptdate='', conscompany='', firstpro='" & azpro & "' where recno='" & rs!recno & "'"
                   cmd.CommandText = dm: Err = 0: cmd.Execute
                End If
             End If
             rs.Close
          End If
       End If
 End Select
 
 '//\\DEBUG GoTo ENDSAVMNFST
 
 'in this section, only fully process (ProTrace & Unix) the Removed pros from a manifest edit
 If l_ne = "EDIT" Then 'editing manifest - delete records for any removed pros
   Pause 0.2
   For k = 1 To 104 'cycle thru original pro list looking for removed pros
     If prolst(k, 0) <> "" And prolst(k, 1) = "0" Then 'if pro was not marked 'keep' then delete the record
       sqs = "delete from trippro where pro='" & prolst(k, 0) & "' and tripnum='" & trpno & "'"
       cmd.CommandText = sqs: Err = 0: cmd.Execute
       If Err <> 0 Then
         dm = "FAIL Remove Previous-Version of Manifest Pro " & prolst(k, 0) & " in 'trippro'" & vbCrLf & _
               vbCrLf & " ** Save Aborted ** Please Inform IT ASAP." & vbCrLf & vbCrLf & Err.Description
         Err = 0: MsgBox dm, , "Manifest" ': prolst(k, 1) = "9" 'mark removal pro as error-state
         GoTo ENDSAVMNFST
       End If
             
       '**************
       ''If typ > 0 Then 'if printing full set or gate pass with shipments specified, then update pros & transfer to Unix
       '**************
         'update probill record (clear all trip references)
         sqs = "update probill set deldate='', delivery='', loadno='', dlvytrip='' where pronumber='" & _
                prolst(k, 0) & "'"
         cmd.CommandText = sqs: Err = 0: cmd.Execute
         If Err <> 0 Then 'probill update failed
           dm = "FAIL Update Previous-Version Pro " & prolst(k, 0) & " in 'probill'" & vbCrLf & vbCrLf & _
                " ** Save Aborted **. Please Inform IT ASAP." & vbCrLf & vbCrLf & Err.Description
           Err = 0: MsgBox dm, , "Manifest" ': prolst(k, 1) = "9"
           GoTo ENDSAVMNFST
         Else 'probill update OK          ''', transfer updated pro to Unix sys
'//\\DEBUG
           If Not correrr Then '01Dec09 - prompt for correspondence entry for any removed pros (removal reason under OSD)
             'remove 'MANIFESTED' status for pros removed from manifest
             'find last 5 status records for shipment
             dm = "select statuscode, recno from stathist where pro='" & prolst(k, 0) & _
                  "' order by  date_ent desc, cast(hr_ent as unsigned) desc, cast(min_ent as unsigned) desc, recno desc limit 5"
             rs.Open dm
             If Not rs.EOF Then
               Do 'cycle thru records (backwards in order they were entered as per 'order by' clause above)
                 If typ = 5 Then
                   If rs(0) <> "S" Then Exit Do 'done when first not 'S' status encountered
                 Else
                   '' 21Nov13 If rs(0) <> "M" Then Exit Do 'done when first not 'M' status encountered
                   Select Case rs(0)
                     Case "M", "Q"
                     Case Else: Exit Do
                   End Select
                 End If
                 sqs = "delete from stathist where recno='" & rs(1) & "'": cmd.CommandText = sqs: Err = 0: cmd.Execute
                 If Err <> 0 Then
                   MsgBox "FAIL Remove 'M'anifested or Loaded 'Q' Status for Removed Shipment: " & prolst(k, 0), , "Manifest": Err = 0
                 End If
                 rs.MoveNext
               Loop Until rs.EOF
             End If
             rs.Close
             If rs.State <> 0 Then rs.Close
             'remove any 'On-Delivery status or timeline events for this trip
             dm = "select recno from stathist where pro='" & prolst(k, 0) & "' and code in ('O','Q') and cons_city='" & Right(trip, 6) & "'"
             rs.Open dm
             If Not rs.EOF Then
               Do
                 dm = "delete from stathist where recno='" & rs(0) & "'": cmd.CommandText = dm: Err = 0: cmd.Execute: rs.MoveNext
               Loop Until rs.EOF
             End If
             rs.Close
             If rs.State <> 0 Then rs.Close
             dm = "select recno from stattimeline where pro='" & prolst(k, 0) & "' and code in ('OUT','LOD') and trip='" & Right(trip, 6) & "'"
             rs.Open dm
             If Not rs.EOF Then
               Do
                 dm = "delete from stattimeline where recno='" & rs(0) & "'": cmd.CommandText = dm: Err = 0: cmd.Execute: rs.MoveNext
               Loop Until rs.EOF
             End If
             rs.Close
                      
             gen1 = prolst(k, 0) 'send pro in global var, gen1 returns with correspondence text
             remvpro.Show 1
             'write a record to table trippro_remv for use by Mendix to remove shipment from manifest on mobile device
             dm = "insert ignore into trippro_remv set pro='" & prolst(k, 0) & "', trip='" & trpno & "', init='" & init & _
                  "', dt='" & Format$(Now, "YYYY-MM-DD Hh:Nn:Ss") & "', active='1'"
             cmd.CommandText = dm: Err = 0: cmd.Execute
             If Err <> 0 Then
               MsgBox "FAIL Write Mobile Remove Record!" & vbCrLf & Err.Description, , "Manifest"
             End If
             acmd.CommandText = "update ndx_counter set ndx=last_insert_id(ndx+1)"
             Err = 0: acmd.Execute: ars.Open "select last_insert_id()"
             If Err = 0 Then ndxv = ars(0)
             If ars.State <> 0 Then ars.Close
             If Err <> 0 Then
               MsgBox "FAIL Get Attach Index!" & vbCrLf & Err.Description, , "Manifest": GoTo SKIPRMVPRO
             End If
             sqs = "insert into filemaster set ndx='" & ndxv & "', pro='" & prolst(k, 0) & _
                   "', filetype='Comment', adddate='" & Format$(Now, "YYYY-MM-DD") & _
                   "', addtime='" & Format$(Now, "Hh:Nn:Ss") & "', comment='" & mySav(gen1) & _
                   "', category='OS&D', init='" & init & "', lastevent='AddCom'"
             acmd.CommandText = sqs: Err = 0: acmd.Execute
             If Err <> 0 Then
               MsgBox "FAIL Save Correspondence!" & vbCrLf & Err.Description, , "Manifest"
             End If
SKIPRMVPRO:
             Err = 0
           End If
         End If
     End If
   Next k
 End If
 
TESTTRIPTRANS:

 '*******************************
 If typ = 5 Then GoTo ENDSAVMNFST 'skip transfers to Unix if storage trailer manifest
 '*******************************

 'write this trip to FTP transfer queue
 dm = trailer
 If typ = 0 Then
   dm2 = "SAVE ONLY"
 Else
   If typ = 4 Then
     If sameday Then dm2 = "AUTODLVR SAMEDAY" Else dm2 = "AUTODLVR"
   Else
     ''If Val(dest) = 90 Then dm2 = "AUTOBTLM" Else
     dm2 = "SAVE GATEPASS"
   End If
 End If
 If Val(trlrapp) > 0 Then dm = dm & "-" & trlrapp
 rs.Open "select recno from trpftpq where tripnum='" & trpno & "'"
 lstq = rs.EOF: rs.Close
 If lstq Then sqs = "insert ignore into trpftpq set " Else sqs = "update trpftpq set "
 sqs = sqs & "trailer='" & dm & "', dt='" & Format$(Now, "YYYY-MM-DD Hh:Nn:Ss") & _
             "', savtyp='" & dm2 & "'"
 If lstq Then sqs = sqs & ", tripnum='" & trpno & "'" Else sqs = sqs & " where tripnum='" & trpno & "'"
 cmd.CommandText = sqs: Err = 0: cmd.Execute
 If Err <> 0 Then
   MsgBox "FAIL Update Trip Transfer Queue!! Pls Advise IT Support. Error Code will be Displayed.", , "Manifest"
   MsgBox Err.Description: Err = 0
 End If
 
 '*******************************
 If typ = 0 Then GoTo ENDSAVMNFST 'transfer trip to Unix only if printing full set and/or gate pass
 '*******************************
 
 'transfer trip header
 dm = trpno: dm2 = TrnTripHdr(dm)

 'transfer trip QC record, display on screen
 arg = TrnTripQC(dm, dm2)
  
ENDSAVMNFST:

 If ars.State <> 0 Then ars.Close
 If aconn.State <> 0 Then aconn.Close
 Set ars = Nothing: Set acmd = Nothing: Set aconn = Nothing

End Function

Private Function TrnTripQC(utrp$, ust$) As Boolean
 Dim dd As Double
 
 If Err <> 0 Then
  MsgBox "1 " & Err.Description
 End If
 
 'open or update data transfer-to-Unix audit record (see T1_Timer for audit tracking)
 rs.Open "select trip from trp_trans where trip='" & utrp & "'"
 If rs.EOF Then
   dm = "insert into trp_trans set trip='" & utrp & "', create_timestamp='" & _
         Format$(Now, "YYYY-MM-DD Hh:Nn:Ss") & "', file_status='" & ust & _
        "', trans_ack_timestamp=''"
 Else
   dm = "update trp_trans set create_timestamp='" & Format$(Now, "YYYY-MM-DD Hh:Nn:Ss") & _
        "', file_status='" & ust & "', trans_ack_timestamp='' where trip='" & utrp & "'"
 End If
 rs.Close
 
 If Err <> 0 Then
   MsgBox "2 " & Err.Description
 End If
 
 cmd.CommandText = dm: Err = 0: cmd.Execute
 If Err <> 0 Then
   MsgBox "INTERNAL ERROR: TRP_TRANS Table Update FAIL! Please Advise IT Support ASAP" & vbCrLf & vbCrLf & "Desc:" & Err.Description
   Err = 0
 Else
   dd = DateValue(Now)
   tlst.AddItem trpno & Chr(9) & CStr(dd) & Format$(TimeValue(Now), ".#0000")
 End If
End Function

Private Function TrnTripHdr$(utrp$)
 TrnTripHdr = "OK"
 Exit Function
End Function

Private Function FixUnitTrailerRun(upre$, fun$, tpre$, ftr$, rn$, ftrun$) As Boolean
 Dim j As Integer
 
 'correct unit for owner codes
 If Trim$(t_uno) <> "" And t_uno <> "SP" And t_uno <> "SDT" Then 'sdt*
   fun = t_uno & unit: upre = t_uno
 Else
   If IsNumeric(unit) = True Then
     unit = CStr(Val(unit))
   Else
     unit = Trim$(Right$(unit, 8))
   End If
   fun = unit: upre = "SP" 'sdt*
 End If
 'correct trailer for owner codes (ftr$) and Unix sys 8-char max fields inc. Run# (trlrun)
 If trlrapp = "" Then rn = "0" Else rn = trlrapp 'run
 If Trim$(t_tro) <> "" And t_tro <> "SP" And t_tro <> "SDT" Then 'non-speedy code 'sdt*
   j = Len(Trim$(t_tro))
   If Left$(trailer, j) = t_tro Then 'owner code prefixes trailer no. entry  '' If Left$(trailer, 3) = t_tro
      ''If Len(trailer) = 3 Then 'trailer entry is code!!
     If trailer = t_tro Then
       ftr = trailer 'use code as trailer entry - bad user practice
     Else
       trailer = Mid$(trailer, j + 1) 'remove owner code from trailer - used eventually for Unix sys ''trailer = Mid$(trailer, 4)
       ftr = t_tro & trailer     'add back on for ProTrace
     End If
   Else
     ftr = t_tro & trailer 'prefix trailer with owner code for ProTrace only
   End If
 Else 'Speedy owner code
   If IsNumeric(trailer) = True Then trailer = CStr(Val(trailer)) 'remove leading zeros
   ftr = trailer 'no owner code prefix on trailer if Speedy
 End If
 If rn = "0" Then
   ftrun = trailer 'specific non-prefixed format for Unix sys
 Else  'run no. entered (2, 3 ... 9)
   'format trailer-run for Unix sys if trailer not > 6 chars already (max 8-char field in Unix sys)
   If Len(trailer) > 6 Then ftrun = trailer Else ftrun = trailer & "-" & rn
 End If
 If t_tro = "" Then tpre = "SP" Else tpre = t_tro 'sdt* write trailer owner code to dedicated field in ProTrace
End Function

Private Function GetNxtTrpNum$()
 On Error Resume Next
 cmd.CommandText = "update tripcntr set tripnumtest = last_insert_id(tripnumtest + 1)"
 Err = 0: cmd.Execute
 If Err <> 0 Then
   MsgBox "FAIL Generate Next Trip Number! Aborting Save . ." & vbCrLf & vbCrLf & Err.Description, , "Manifest"
   Err = 0: Exit Function
 End If
 
 If rs.State <> 0 Then rs.Close
 rs.Open "select last_insert_id()" 'complete mySQL method that ensures only this user gets this next-number
 If rs.EOF Or Err <> 0 Then
   rs.Close
   MsgBox "FAIL Retrieve Next-Number Trip Index! Aborting Save . ." & vbCrLf & vbCrLf & Err.Description, , "Manifest"
   Err = 0: Exit Function
 End If
 GetNxtTrpNum = rs(0) 'next-number trip# retrieved
 rs.Close
End Function

Private Function GetNxtEmptyCN$()
 On Error Resume Next
 cmd.CommandText = "update tripcntr set counter = last_insert_id(counter + 1)"
 Err = 0: cmd.Execute
 If Err <> 0 Then
   MsgBox "FAIL Generate Next Empty Control#! Aborting Print . . " & vbCrLf & vbCrLf & Err.Description, , "Manifest"
   Err = 0: Exit Function
 End If
 
 If rs.State <> 0 Then rs.Close
 rs.Open "select last_insert_id()" 'complete mySQL method that ensures only this user gets this next-number
 If rs.EOF Or Err <> 0 Then
   rs.Close
   MsgBox "FAIL Retrieve Next-Number Trip Index! Aborting Save . ." & vbCrLf & vbCrLf & Err.Description, , "Manifest"
   Err = 0: Exit Function
 End If
 GetNxtEmptyCN = rs(0) 'next-number trip# retrieved
 rs.Close
End Function


Private Function SetLMStatus$()

 '19May20(DNY) LH moves in USA between US Partner terminals
 Select Case Val(orig)
   Case 45, 46, 47, 52, 53, 54, 55, 57 To 61, 63, 67  '19May2020 -> PYLE 45 to 47, EXLA 52 to 54, YRC 55, HMES 57 to 61, WARD 63, SEFL 67
     Select Case Val(dest)
       Case 45, 46, 47, 52, 53, 54, 55, 57 To 61, 63, 67
         If Val(orig) = Val(dest) Then SetLMStatus = "M" Else SetLMStatus = "L"
         Exit Function
     End Select
 End Select
 
 If Val(orig) = Val(dest) Then
   If Val(orig) = 31 Then SetLMStatus = "S" Else SetLMStatus = "M"  'all city, all cust pickup manifests
 Else
   Select Case Val(orig)
     Case 37, 38, 71, 72, 73 'all dock pickup/transfer codes (cust or agent)
       SetLMStatus = "M"
     Case Else
       SetLMStatus = "L"
   End Select
 End If
End Function

Private Function GetTripPrefix$()
 'build trip# prefix code from orig/dest codes
 
 '19May20(DNY) LH moves in USA between US Partner terminals
 Select Case Val(orig)
   Case 45, 46, 47, 51, 52, 53, 54, 55, 57 To 61, 63, 67 '19May2020 -> PYLE 45 to 47, EXLA 52 to 54, YRC 55, HMES 57 to 61, WARD 63, SEFL 67
     Select Case Val(dest)
       Case 45, 46, 47, 51, 52, 53, 54, 55, 57 To 61, 63, 67
         GetTripPrefix = lhcod(Val(orig)): Exit Function 'use origin terminal code for US-US linehaul
     End Select
 End Select
 
 If lhcod(Val(orig)) = "" Then 'no tripcode for orig code entry
   If lhcod(Val(dest)) = "" Then
     GetTripPrefix = "Z"
   Else
     Select Case Val(dest)
       ''''27Sep16 Case 54, 55, 60 To 63, 65
       Case 45, 46, 51 To 54, 55, 57 To 65
         GetTripPrefix = lhcod(Val(dest)) 'use dest code only if US carrier border crossing
       Case Else
         GetTripPrefix = "Z"
     End Select
   End If
 Else
   If Val(orig) < 19 Then  'And Val(orig) <> 8 Then 'origin is a Speedy terminal
     If orig = dest Then 'local trip
       GetTripPrefix = "X" 'all locals start with 'X'
     Else
       If Val(dest) < 19 Then 'And Val(orig) <> 8 Then 'terminal-to-terminal linehaul
         GetTripPrefix = lhcod(Val(orig)) 'use origin terminal code
       Else 'terminal-to-US or agent or other
         Select Case Val(dest)
           ''27Sep16 Case 54, 55, 60, 61, 63, 65 'US carrier (Holland, ODFL, NewPenn, etc.) border crossing
           Case 45, 46, 51 To 54, 55, 57 To 65 'US carrier (Holland) or SZTG border crossing
             If lhcod(Val(dest)) = "" Then GetTripPrefix = lhcod(Val(orig)) Else GetTripPrefix = lhcod(Val(dest)) 'use destination code
           Case Else 'agent, customer drop, SCM Cornwall, etc.
             GetTripPrefix = lhcod(Val(orig)) 'use origin terminal
         End Select
       End If
     End If
   Else 'origin is US border crossing or agent
     Select Case Val(dest)
       Case 31: GetTripPrefix = "S" 'storage trailer
       ''27Sep16 Case 54, 55, 60 To 63, 65 'US carrier border crossing
       Case 45, 46, 51 To 54, 55, 57 To 65 'US carrier border crossing
         If lhcod(Val(dest)) = "" Then GetTripPrefix = lhcod(Val(orig)) Else GetTripPrefix = lhcod(Val(dest))
       Case Else 'agent, customer drop, SCM Cornwall, etc.
         GetTripPrefix = lhcod(Val(orig))
     End Select
   End If
 End If
End Function

Private Function WrProUnix(pr$) As Boolean 'write probill update transfer records to Unix system
 Exit Function
End Function


Private Sub c_de_Click()
 Dim j As Integer
 If Val(dest) = 0 Then
   lst_de.ListIndex = -1
 Else
   For j = 0 To lst_de.ListCount - 1
     If Val(dest) = lst_de.ItemData(j) Then
       lst_de.ListIndex = j: Exit For
     End If
   Next j
 End If
 lst_de.Visible = True: fr1.Enabled = False: lst_de.SetFocus
End Sub


Private Sub c_pegp_Click()  '16Sep22(DNY)
 Dim i As Integer, j As Integer, procnt As Integer
 Dim apptcnt As Integer, apptfirst As Integer, apptlast As Integer, js As Single
 Dim oy As Single
 Dim dm As String, dm2 As String, upre As String, tpre As String, tplate As String
 Dim unp As String, trp As String, stcity As String, encity As String
 Dim apptfirsttime As String, apptlasttime As String
 Dim appta() As String
 Dim specreq As Boolean
 
 'access rights - y/n, who? - NOT enabled for now
 'prechecks: must have unit, orig, dest, no pros
 'control no.  000000-099999  table emptygatepass - controlnumber integer maxlength 6
 
 If Trim$(unit) = "" Then
   MsgBox "Unit/Tractor Must be Specified. Aborting Print . .": Exit Sub
 End If
 If orig = "" Or dest = "" Then
   MsgBox "Origin & Destination Must be Specified. Aborting Print . .": Exit Sub
 End If
 For j = 0 To 204 Step 2
    If Trim$(g2.TextMatrix(j, 17)) <> "" Then
      MsgBox "NOT Empty! Pro# entry detected near line " & CStr((j + 2) / 2) & ". Aborting Print . .": Exit Sub
    End If
 Next j
 
 cn = GetNxtEmptyCN
 If Val(cn) = 0 Then Exit Sub
 cn = Format$(Val(cn), "000000")
 
 If gateprn = "GatePassPrinter" Then
   Set Printer = ogateprn: SetDefaultPrinter "GatePassPrinter"
 ElseIf inbondprn = "InlandInbondPrinter" Then
   Set Printer = oinbondprn: SetDefaultPrinter "InlandInbondPrinter"
 End If
 
 Printer.FontTransparent = True
 Printer.Orientation = 1 'force protrait mode
 Printer.ScaleMode = 5 'inches
 'w 12510 h 9375
 
 Printer.PaintPicture pic1, 0#, 0.05, 2.05, 0.6154
 Printer.DrawWidth = 12
 Printer.Line (0.01, 0.75)-(8#, 3.75), , B 'main box
 Printer.DrawWidth = 3
 Printer.Line (5.25, 0.75)-Step(0, 3#) 'vert
 For j = 1 To 5
   Printer.Line (0.01, 0.75 + (j * 0.5))-Step(5.25, 0#)
   If j = 1 Then Printer.Line (5.25, 0.75 + (j * 0.5))-Step(2.75, 0#)
   If j = 5 Then Printer.Line (5.25, 0.75 + (j * 0.5))-Step(2.75, 0#) ')(
 Next j
 Printer.Line (5.25, 1.75)-Step(2.75, 0#)
 Printer.Line (5.25, 2.75)-Step(2.75, 0#)
 
 
 Printer.FontName = "Arial": Printer.FontBold = True: Printer.FontSize = 18
 lprnt 3!, 0.42, "GATE PASS - EMPTY"
 
 Printer.FontSize = 12
 lprnt 0.2, 0.75 + 0.17, "TRUCK CO."
 lprnt 0.2, 1.25 + 0.17, "TRUCK No."
 lprnt 0.2, 1.75 + 0.17, "TRAILER CO."
 lprnt 0.2, 2.25 + 0.17, "TRAILER No."
 lprnt 0.2, 2.75 + 0.17, "DRIVERS SIGNATURE"
 lprnt 0.2, 3.25 + 0.17, "AUTHORIZED BY"
 
 Printer.FontSize = 11
 cprnt 5.25 + (2.75 / 2#), 0.75, "DATE"
 cprnt 5.25 + (2.75 / 2#), 1.25, "LOADER"
 cprnt 5.25 + (2.75 / 2#), 1.75, "ADDITIONAL INFO"
 cprnt 5.25 + (2.75 / 2#), 2.75, "SEAL NUMBER"
 cprnt 5.25 + (2.75 / 2#), 3.25, "TRAILER PLATE"
 
 Printer.FontSize = 18
 'unit
 unp = unit
 If IsNumeric(unit) = True Then
   Select Case Trim$(t_uno)
     Case "", "SP", "SDT"  'sdt*
       If InStr(t_un, " ST ") > 0 Then
         unp = "ST " & unp
       ElseIf InStr(t_un, " CUBE ") > 0 Then
         unp = "CVAN " & dm
       End If
     Case Else: unp = t_uno & unp
   End Select
 End If
 lprnt 2.75, 1.25 + 0.11, unp
  
 'trailer
 trp = trailer: dm2 = ""
 Select Case Trim$(t_tro)
   Case "", "SP", "SDT", "HOL", "NPM", "ODF", "YRC", "YRT"  'sdt*
   Case Else: trp = t_tro & trp
 End Select

 If trp = "" Then trp = "Bobtail"
 lprnt 2.75, 2.25 + 0.11, trp
 
 'attempt to discern unit/trailer company
 upre = "": tpre = ""
 Select Case Trim$(t_uno)
   Case "", "SP", "SDT": upre = "Speedy Transport"  'sdt*
   Case Else
     For j = 0 To 99
       If uownr(j, 0) = "" Then Exit For
       If uownr(j, 0) = t_uno Then
         upre = uownr(j, 1): Exit For
       End If
     Next j
 End Select
 
 Select Case Trim$(t_tro)
   Case "", "SP", "SDT" 'sdt*
     tpre = "Speedy Transport"
   Case Else
      For j = 0 To 99
       If tsc(j, 0) = "" Then Exit For
       If tsc(j, 0) = t_tro Then
         tpre = tsc(j, 1): Exit For
       End If
     Next j
 End Select
 
 Printer.FontSize = 16
 If unit <> "" And upre <> "" Then lprnt 2.75, 0.75 + 0.11, upre
 If tpre <> "" Then lprnt 2.75, 1.75 + 0.11, tpre
 
 cprnt 5.25 + (2.75 / 2#), 3.15 - 0.2275, "EMPTY"
 
 If Trim$(descr) <> "" Then
   Printer.FontSize = 10
   cprnt 5.25 + (2.75 / 2#), 2.15, Trim$(descr)
   Printer.FontSize = 16
 End If
 
 'authorized by: gate pass issuer 07Apr16(DNY)
 lprnt 1.75, 3.37, gpissuer
 
 If Trim$(mnfst!l_loadby) <> "" Then
   Printer.FontSize = 13: js = 0.22
   Do
     Printer.FontSize = Printer.FontSize - 1: js = js + 0.01
     If Printer.TextWidth(l_loadby) <= 2.73 Then Exit Do
   Loop Until Printer.FontSize = 9
   cprnt 5.25 + (2.75 / 2#), 1.25 + js, l_loadby
   Printer.FontSize = 16
 End If
 
 cprnt 5.25 + (2.75 / 2#), 0.94, Format$(mdte, "DD-MMM-YYYY")

 cd = Format$(Now, "DDMMMYYYY")
 
 Printer.FontSize = 9
 cprnt 5.25 + (2.75 / 2#), 0.55, UCase$(cn & "-" & cd & "-" & init) 'print control no.

 lprnt 1.9375, 0.5 + oy, l_or  'orig
 lprnt 2.28125, 0.5 + oy, l_de 'dest

 dm = "insert ignore into emptygatepass set controlnumber='" & cn & "', printdatetime='" & Format$(Now, "YYYY-MM-DD Hh:Nn:Ss") & _
      "', user='" & mySav(gpissuer) & "', inits='" & uinits & "', tractor='" & mySav(unp) & "', trailer='" & trp & _
      "', descrip='" & mySav(Trim$(descr)) & "', orig='" & l_or & "', dest='" & l_de & "'"
 cmd.CommandText = dm: cmd.Execute
      

GoTo EPGPEGP

EPGPEGP:
 Printer.EndDoc
 Printer.Orientation = 2 'reset default printer to landscape
End Sub

Private Sub c_prntbols_Click() 'print ALL BOL's in manifest
 Dim i As Integer, j As Integer
 Dim cnt As Single, ecnt As Single
 Dim iasp As Single, pasp As Single, objw As Single
 Dim dm As String, Msg As String
 Dim b As Boolean
 Dim bpro() As String
 
 For j = 0 To 206 Step 2 'cycle thru shipment list looking for missing BOL's
   If g2.TextMatrix(j, 17) <> "" Then
     cnt = cnt + 1#
     If g2.TextMatrix(j, 11) = "" Then
       ecnt = ecnt + 1#: Msg = Msg & g2.TextMatrix(j, 17) & " "
     End If
   End If
 Next j
 
 If ecnt = cnt Then 'no BOL's
   MsgBox "No Manifest Shipments Currently Have BOL Scans. Re-Load Manifest if BOL's were scanned since previous load.", , "Manifest"
   Exit Sub
 End If
 
 If Msg <> "" Then
   If ecnt > 5 Then
     dm = "BOL's Missing for " & CStr(CInt(ecnt)) & " Shipments. Do you Wish to Continue Printing *ALL* Manifest BOL's?"
   Else
     dm = "Missing BOL's for Shipments: " & Trim$(Msg) & ". Do you Wish to Continue Printing *ALL* Manifest BOL's?"
   End If
 Else
   dm = "Ready to Print " & CStr(CInt(cnt)) & " Shipment BOL's. Continue?"
 End If
 If MsgBox(dm, vbDefaultButton2 + vbQuestion + vbYesNo, "Print *ALL* Manifest Shipment BOL's") = vbNo Then Exit Sub
   
 frw.Visible = True: frw.Refresh
 '1. get all images as files to local temp folder
 Err = 0
 If fo.FolderExists(sdir & "mprntmp") Then fo.DeleteFolder sdir & "mprntmp"
 fo.CreateFolder sdir & "mprntmp"
 If Err <> 0 Then
   MsgBox "FAIL Create Temporary BOL Image Print Folder!  Aborting . . .", , "ProTrace Manifest Print"
   Exit Sub
 End If
 i = -1
 For j = 0 To 206 Step 2 'cycle thru
   If g2.TextMatrix(j, 17) <> "" And g2.TextMatrix(j, 11) <> "" Then
     If MGetTempBlobImage(g2.TextMatrix(j, 11), g2.TextMatrix(j, 17), g2.TextMatrix(j, 15), "") Then
       b = True: Exit For
     End If
     i = i + 1
     ReDim Preserve bpro(0 To i) 'extend holding array by 1
     bpro(i) = g2.TextMatrix(j, 17) 'add BOL print pro#
   End If
 Next j
 If b Then GoTo ENDCPBS
 If i = -1 Then GoTo ENDCPBS
 '2. setup printer scaling, etc.
 Printer.ScaleMode = vbTwips  'default scaling
 pwinch = Printer.ScaleWidth / 1440! '1440 twips/inch scaling (global var)
 Printer.ScaleMode = vbInches
 '3. print image files using PictPlus object
 For j = 0 To UBound(bpro) 'cycle thru BOL print pros
   obj.filename = sdir & "mprntmp\" & bpro(j) & ".tif"
   obj.FileFormat = "": obj.Page = 1
   result = obj.LoadFile
   If InStr(result, ",") > 0 Then
     prop = Split(result, ",")
   Else
     MsgBox "FAIL: Retrieve Image from File: " & sdir & "mprntmp\" & bpro(j) & ".tif", , "Manifest": Exit For
   End If
   size_w = prop(2): size_h = prop(3): pasp = Printer.ScaleHeight / Printer.ScaleWidth
   iasp = size_h / size_w
   If iasp > pasp Then objw = pwinch * (pasp / iasp) - 0.01 Else objw = (CSng(Printer.Width) / 1440!) - 0.01
   PrntMBOL objw
 Next j

ENDCPBS:
  frw.Visible = False
End Sub

Private Sub c_prnttom_Click()
 Dim cd As String, pndx As String, BC As String
 Dim lhdl As Integer, cucnt As Integer
 Dim cube As Long

 If Not dockpkup Then
   noprotrp = True: savprnt = True
   c_new_Click
   If Not savprnt Then Exit Sub 'error during save, abort print
 End If
 prnttom = False
 savprnt = False
 
 'print Gate Pass
 If Not prnavail Then
   MsgBox "No Default Printer for this Computer!", , "Print Gate Pass"
   Exit Sub
 End If

 If Not dockpkup Then
   If l_ne = "NEW" Then
     MsgBox "Manifest Must Be SAVED and Displayed in EDIT Mode to Print. Aborting Print . .", , "Manifest"
     Exit Sub
   End If
   If Len(trip) <> 7 Then
     MsgBox "Invalid Trip No.! Please Contact IT Support. Aborting . .", , "Manifest"
     Exit Sub
   End If
 End If

TESTGP345:
 
 'count populated rows, get total weight, total cube
 totpros = 0: totwt = 0
 For j = 0 To 206 Step 2
  If g2.TextMatrix(j, 17) <> "" And g2.TextMatrix(j, 3) <> "" Then
    totpros = totpros + 1: totwt = totwt + Val(g2.TextMatrix(j, 5))
    If Val(g2.TextMatrix(j, 14)) > 0 Then
      cucnt = cucnt + 1: cube = cube + Val(g2.TextMatrix(j, 14))
    End If
  End If
 Next j
 twt = CStr(totwt): tsh = CStr(totpros): tcu = CStr(cube)
 tcucnt = CStr(cucnt) & "/" & CStr(totpros)
 
'@@disable if..endif below for gatepass test
 If Not dockpkup And totpros > 0 Then
   If t_tom.Visible Then Exit Sub
 End If

 'ready to print, disable controls while printing
 frw.Visible = True: Refresh: DoEvents
 MousePointer = 11
 fr1.Enabled = False: frg1.Enabled = False
 c_new.Enabled = False: c_prnt.Enabled = False
 c_clr.Enabled = False: c_fgo.Enabled = False
 
 'create ACE style barcode for manifest & gate pass
 BC = Format$(mdte, "YYYYMMDD") & "BUFM" & Right$(trip, 6)
 DrawBarCode BC, 0
 
 'get next unique print control#
 cmd.CommandText = "update tripcntr set prntctrl = last_insert_id(prntctrl + 1)"
 Err = 0: b = False
 cmd.Execute
 If Err <> 0 Then
   MsgBox "FAIL Generate Next Print Control Number!" & vbCrLf & vbCrLf & Err.Description, , "Manifest"
   Err = 0: b = True
 End If
 If Not b Then
   rs.Open "select last_insert_id()" 'complete mySQL method that ensures only this user gets this next-number
   If rs.EOF Or Err <> 0 Then
     rs.Close
     MsgBox "FAIL Retrieve Next-Number Print Control Index!" & vbCrLf & vbCrLf & Err.Description, , "Manifest"
     Err = 0: b = True
   End If
   pndx = rs(0) 'next-number trip# retrieved
   rs.Close
 End If
 If b Then pndx = "Z" & CStr(Month(Now)) & CStr(Day(Now)) & CStr(Hour(Now)) & CStr(Minute(Now))
 
 cd = Format$(Now, "DDMMMYY-HhNn") 'print timestamp
 
 'determine whether delivery or linehaul manifest
 If Val(orig) = Val(dest) Then
   lhdl = 0
 Else
   Select Case Val(dest)
     Case 37, 38, 71, 72, 73: lhdl = 0
     Case Else: lhdl = 1
   End Select
 End If
 
 PrntGatePass 2, pndx, cd, BC, lhdl '0=not linehaul
 
 frw.Visible = False
 MousePointer = 0
 fr1.Enabled = True: frg1.Enabled = True
 c_new.Enabled = True: c_prnt.Enabled = True
 c_clr.Enabled = True: c_fgo.Enabled = True
 If Not dockpkup Then c_prnttom.SetFocus

End Sub

Private Sub c_strip_Click()
 Dim dm As String
 If init = "" Then Exit Sub
 dm = "C:\ProTrace\StagingFolders\StripManifest\stripmanifest.exe " & init & "|"
 If Shell(dm, vbNormalFocus) = 0 Then
   MsgBox "FAIL Launch Strip Manifest!" & vbCrLf & vbCrLf & Err.Description
 End If
End Sub

Private Sub c_svc_Click() 'show trailer servicing report
 svcrpt.Show 1
End Sub

Private Sub c_uno_Click()
 If lst_tro.Visible Then lst_tro.Visible = False
 If lst_uno.Visible Then lst_uno.Visible = False Else lst_uno.Visible = True
 On Error Resume Next
 lst_uno.SetFocus
End Sub
Private Sub c_tro_Click()
 If lst_htr.Visible Then Exit Sub
 If lst_uno.Visible Then lst_uno.Visible = False
 If lst_tro.Visible Then lst_tro.Visible = False Else lst_tro.Visible = True
 On Error Resume Next
 lst_tro.SetFocus
End Sub


Private Sub c_zstat_Click()
 Dim dm As String
 Dim j As Integer
 Dim errb As Object

 If l_ne <> "EDIT" Then Exit Sub
 If trailer = "" Then
   MsgBox "Missing Trailer Entry. Aborting . . .", , "Manifest": Exit Sub
 End If
 gen1 = ""
 zstatfrm.Show 1
 If gen1 = "" Then Exit Sub ' if applied, should return trailer drop date (status date) in mySQL format
 If IsDate(gen1) = False Then Exit Sub
 'cycle thru pros, add 'Z' status to status history, transfer to legacy
 On Error Resume Next
 For j = 0 To 206 Step 2
   If g2.TextMatrix(j, 17) <> "" And g2.TextMatrix(j, 3) <> "" Then 'entry on this line in grid
     dm = "insert ignore into stathist set pro='" & g2.TextMatrix(j, 17) & "', hr_ent='" & _
           Format$(Now, "Hh") & "', min_ent='" & Format$(Now, "Nn") & "', statuscode='Z', date_ent='" & _
           Format$(Now, "YYYY-MM-DD") & "', date_status='" & gen1 & "', trailer='" & trailer & _
          "', origin='" & orig & "', dest='" & dest & "'"
     cmd.CommandText = dm: Err = 0: cmd.Execute
     If Err <> 0 Then Set errb = Err
     dm = "update probill set funds='Z' where pronumber='" & g2.TextMatrix(j, 17) & "'"
     cmd.CommandText = dm: Err = 0: cmd.Execute
     If Err <> 0 Then Set errb = Err
   End If
 Next j
 If Not (errb Is Nothing) Then
   MsgBox "** WARNING ** Error(s) encountered writing 'Z' Status!  Last Error:" & vbCrLf & vbCrLf & _
          errb.Description
 End If
 Set errb = Nothing
End Sub


Private Sub g2pop6_Click() 'set/un-set QuikX origin shipment in 1-62 NPME southbound manifest
 Dim dm As String, dm2 As String, dm3 As String, ogen2 As String
 Dim wasset As Boolean
 Dim k As Integer

 gen3 = g2.TextMatrix(curow, 17)
 If Val(gen3) = 0 Then Exit Sub
 gen1 = "3": dm = g2.TextMatrix(curow + 1, 1)
 If g2.TextMatrix(curow, 20) Like "QXTI*" Then
   gen2 = Mid$(g2.TextMatrix(curow, 20), 5): wasset = True
   dm2 = "** QXTINPME-" & gen2 & " " 'current account codes row QXTI advance pro display text
 Else
   gen2 = "": wasset = False: dm2 = ""
 End If
 ogen2 = gen2
 If gen2 = "" Then wset = False Else wset = True
 entcnt.Show 1
 If gen2 = "-999" Then Exit Sub
 If ogen2 = gen2 Then Exit Sub
 If gen2 = "" Then 'un-set
   g2.TextMatrix(curow, 20) = ""
   If wasset Then
     dm = Replace(dm, dm2, "") 'remove display string from account codes row text string
     If Trim$(dm) = "**" Then dm = " "
   Else
     Exit Sub
   End If
 Else 'set or edit QXTI advance pro#
   g2.TextMatrix(curow, 20) = "QXTI" & gen2
   If wasset Then 'was previously set - update QuikX pronumber (advance pro)
     dm3 = "** QXTINPME-" & gen2 & " " 'new display text
     dm = Replace(dm, dm2, dm3)    'replace old display text with new
     If Trim$(dm) = "**" Then dm = " "
    Else 'was not set previously, add QXTI advance pro display text to start of account codes text string
     dm2 = "** QXTINPME-" & gen2 & " "
     If Trim$(dm) = "" Then dm = dm2 & "**" Else dm = dm2 & dm
   End If
 End If
 For k = 2 To 9 '{@} was 1 to 9  write this identical string into each 'merge' cell to trigger the merger effect across them
   g2.TextMatrix(curow + 1, k) = dm
 Next k
End Sub

Private Sub g2pop7_Click()
 gen1 = g2.TextMatrix(curow, 17)
 dockchk.Show 1
End Sub

Private Sub g2pop8_Click()
 If g2pro.Visible And Trim$(g2pro) <> "" Then
   dm = Trim$(g2pro)
   Clipboard.Clear: Clipboard.SetText dm
   g2procpy = "": g2procpy = dm
 End If
End Sub

Private Sub g2pop9_Click()
 Dim dm As String, dm2 As String, spro As String

 gi = -555 'print single shipment (current row)  ' 999=print all shipments
 gs1 = g2.TextMatrix(curow, 17)
 
 prndr.Show 1
 If gi = -555 Then Exit Sub
 Refresh
 
 If rs.State <> 0 Then rs.Close
 On Error Resume Next
 spro = g2.TextMatrix(newrow, 17)
   
 'print DR - whse/orig/both copies
 dm2 = ""
 Select Case gs3
   Case "10" 'orig only
     idno = PrntDR(spro, "ORIG", "P") 'errors displayed in function
     If idno = "FATAL" Then GoTo ENDCPRNTDR
     dm2 = ", orig='1'"
   Case "01" 'whse only
     idno = PrntDR(spro, "WHSE", "P") 'errors displayed in function
     If idno = "FATAL" Then GoTo ENDCPRNTDR
     dm2 = ", whse='1'"
   Case "11" 'both - print orig first then set indicator to whse and print again with manifest as separator
     idno = PrntDR(spro, "ORIG", "P") 'errors displayed in function
     If idno = "FATAL" Then GoTo ENDCPRNTDR
     idno = PrntDR(spro, "WHSE", "P") 'errors displayed in function
     If idno = "FATAL" Then GoTo ENDCPRNTDR
     dm2 = ", orig='1', whse='1'"
 End Select

 '=+=+=+=+ program-specific +=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+
 'add pro print record to history
 dm = "insert ignore into drprnhist set date='" & Format$(Now, "YYYY-MM-DD") & "', time='" & _
       Format$(Now, "Hh:Nn:Ss") & "', tripnum='" & Right$(trip, 6) & "', init='" & init & _
       "', pro='" & spro & "', edipro='" & edipro & "', idno='" & idno & "', src='s'"
 '=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=
 
 If gi > 0 Then dm = dm & ", reason='" & mySav(gs4) & "'"
 If dm2 <> "" Then dm = dm & dm2
 cmd.CommandText = dm: Err = 0: cmd.Execute
 If Err <> 0 Then
   dm2 = "ERROR: Fail Write PrintHistory Record! Aborting Print Job . . ." & vbCrLf & Err.Description
   MsgBox dm2, , "DR Print": GoTo ENDCPRNTDR
 End If
ENDCPRNTDR:
End Sub


Private Sub i_hlp_Click()
 hlpfrm.Show 1
End Sub

Private Sub l_hlp_Click(x%)
 hlpfrm.Show 1
End Sub
Private Sub lst_uno_DblClick()
 lst_uno.Visible = False
 t_uno = Replace(Trim$(Left$(lst_uno.Text, 3)), vbTab, "")
 If t_un.Visible Then t_un.SetFocus
End Sub
Private Sub lst_tro_DblClick()
 c_canchtr_Click
 t_tro = Trim$(Left$(lst_tro.Text, InStr(lst_tro.Text, Chr(9)) - 1))
 On Error Resume Next
 If t_tr.Visible Then t_tr.SetFocus Else t_tro.SetFocus
 lst_tro.Visible = False: Refresh
End Sub
Private Sub lst_uno_KeyDown(KeyCode%, Shf%)
 Select Case KeyCode
   Case vbKeyEscape
     lst_uno.Visible = False
     If t_uno.Visible Then t_uno.SetFocus
   Case vbKeyReturn, vbKeyExecute: lst_uno_DblClick
 End Select
End Sub
Private Sub lst_tro_KeyDown(KeyCode%, Shf%)
 Select Case KeyCode
   Case vbKeyEscape
     lst_tro.Visible = False
     If t_tro.Visible Then t_tro.SetFocus
   Case vbKeyReturn, vbKeyExecute: lst_tro_DblClick
 End Select
End Sub

Private Sub t_uno_GotFocus()
 t_uno.BackColor = &HFFFFFF
End Sub
Private Sub t_uno_LostFocus()
 t_uno.BackColor = &HD6E6E6
 Select Case t_uno
   Case "", "SP", "OTH", "HOL", "ODF", "NPM"
   Case Else: t_uno = "OTH"
 End Select
 On Error Resume Next
 t_un.SetFocus
End Sub
Private Sub t_tro_GotFocus()
 t_tro.BackColor = &HFFFFFF
End Sub
Private Sub t_tro_LostFocus()
 t_tro.BackColor = &HD6E6E6
 If t_tr.Visible Then t_tr.SetFocus
End Sub
Private Sub t_uno_KeyDown(KeyCode%, Shf%)
 Select Case KeyCode
   Case vbKeyEscape: t_uno = ""
   Case vbKeyDown, vbKeyReturn, vbKeyExecute
     If t_un.Visible Then t_un.SetFocus
 End Select
End Sub
Private Sub t_tro_KeyDown(KeyCode%, Shf%)
 Select Case KeyCode
   Case vbKeyEscape: t_tro = ""
   Case vbKeyDown, vbKeyReturn, vbKeyExecute
     If t_tr.Visible Then t_tr.SetFocus
 End Select
End Sub
Private Sub t_uno_KeyPress(KeyAscii%)
 Select Case KeyAscii
   Case 8, 65 To 90 'bkspace, A - Z
   Case 97 To 122: KeyAscii = KeyAscii - 32 'uppercase a - z
   Case Else: KeyAscii = 0 'all other keys filtered out
 End Select
End Sub
Private Sub t_tro_KeyPress(KeyAscii%)
 Select Case KeyAscii
   Case 8, 65 To 90 'bkspace, A - Z
   Case 97 To 122: KeyAscii = KeyAscii - 32 'uppercase a - z
   Case Else: KeyAscii = 0 'all other keys filtered out
 End Select
End Sub

Private Sub g2pop3_Click() '<F12> view BOL for shipment
 Dim dm As String
 If bolfrm Then
   If g2.TextMatrix(curow, 17) = pview!t_s(1) Then Exit Sub
 End If
 Select Case g2.TextMatrix(curow, 15)
   Case "1"
     '             0     1        2      3     4        5
     dm = "select ndx,inv_unit,pro_date,pgno,accnt,scandatetime from img_sfmaster where ndx='" & g2.TextMatrix(curow, 11) & "'"
   Case "2"
     '             0     1        2      3     4        5
     dm = "select ndx,inv_unit,pro_date,pgno,accnt,scandatetime from img_stpc where ndx='" & g2.TextMatrix(curow, 11) & "'"
 End Select
 
 If ir.State <> 0 Then ir.Close
 If bolfrm Then Unload pview
 ir.Open dm
 If ir.EOF Then
   ir.Close
   MsgBox "BOL Not Found for Pro: " & g2.TextMatrix(curow, 17), , "Manifest"
   If g2pro.Visible Then g2pro.SetFocus
   Exit Sub
 End If
 Load pview  ''ir' closed in pview
 If g2pro.Visible Then g2pro.SetFocus
 If bolfrm Then pview.ZOrder
End Sub

Private Sub g2pop4_Click() 'print BOL for selected row
 Dim iasp As Single, pasp As Single, objw As Single
 If g2.TextMatrix(curow, 11) = "" Then Exit Sub
 On Error Resume Next
 Err = 0
 'create/clear local folder to temporarily store image files
 If fo.FolderExists(sdir & "mprntmp") Then fo.DeleteFolder sdir & "mprntmp"
 fo.CreateFolder sdir & "mprntmp"
 If Err <> 0 Then
   MsgBox "FAIL Create Temporary BOL Image Print Folder!  Aborting . . .", , "ProTrace Manifest Print"
   Exit Sub
 End If
 If MGetTempBlobImage(g2.TextMatrix(curow, 11), g2.TextMatrix(curow, 17), g2.TextMatrix(curow, 15), "") Then Exit Sub
 'print image from file using PictPlus object
 Printer.ScaleMode = vbTwips  'default scaling
 pwinch = Printer.ScaleWidth / 1440! '1440 twips/inch scaling (global var)
 Printer.ScaleMode = vbInches
 'get image into PictPlus object for printing
 obj.filename = sdir & "mprntmp\" & g2.TextMatrix(curow, 17) & ".tif"
 obj.FileFormat = "": obj.Page = 1
 result = obj.LoadFile 'call LoadFile function
 If InStr(result, ",") > 0 Then 'if loaded OK then get image properties for page scaling
   prop = Split(result, ",") 'create array of variant type and read in values
 Else
   MsgBox "FAIL: Retrieve Image from File.", , "Manifest": Exit Sub
 End If
 size_w = prop(2): size_h = prop(3) 'extract image width & height (global vars)
 'get printer max. printable-area size aspect ratio (NOTE: this is NOT pagesize e.g. 8.5" x 11")
 pasp = Printer.ScaleHeight / Printer.ScaleWidth
 iasp = size_h / size_w 'w/h globals previously set along with 'pwinch'
 If iasp > pasp Then 'image height-to-height aspect ratio is greater than printable-area aspect ratio
   'if print width is not reduced then image will fit max printable width BUT extend past max printable height
   '   which causes print error; therefore must reduce print width until image fits into printer page height
   objw = pwinch * (pasp / iasp) - 0.01 'slightly less to account for rounding off errors from printer
 Else
   objw = (CSng(Printer.Width) / 1440!) - 0.01 'slightly less "    "     "      "    "      "     "
 End If
 PrntMBOL objw 'print the image passing printer page scalewidth
End Sub

Private Sub g2pop5_Click() 'adjust total wt for pro in selected row
 Dim dm As String, owt As String
 If g2.TextMatrix(curow, 17) = "" Then Exit Sub
 gen1 = "2": gen2 = Replace(g2.TextMatrix(curow, 5), Chr(13), ""): owt = gen2: gen3 = g2.TextMatrix(curow, 17)
 '
 entcnt.Show 1
 '
 If Val(gen2) <= 0 Or Val(gen2) = Val(owt) Then Exit Sub
 rs.Open "select hazmat_l4, totalwt from probill where pronumber='" & g2.TextMatrix(curow, 17) & "'"
 If rs.EOF Then
   rs.Close: MsgBox "Internal Error: Probill NF! Aborting Save . . .", , "Manifest": Exit Sub
 End If
 If Val(rs(1)) <> Val(owt) Then
   rs.Close: MsgBox "Internal Error: TotalWt Mismatch! Aborting Save . . .", , "Manifest": Exit Sub
 End If
 dm = "=E(" & Left$(init, 1) & Right$(init, 1) & ") " & rs(0)
 rs.Close
 dm = "update probill set totalwt='" & gen2 & "', hazmat_l4='" & dm & _
      "' where pronumber='" & g2.TextMatrix(curow, 17) & "'"
 On Error Resume Next
 cmd.CommandText = dm: Err = 0: cmd.Execute
 If Err <> 0 Then
   MsgBox "FAIL Update TotalWt in Probill!", , "Manifest": Exit Sub
 End If
 g2.TextMatrix(curow, 5) = Chr(13) & gen2 & Chr(13)
End Sub

Private Sub lst_de_DblClick()
 dest = CStr(lst_de.ItemData(lst_de.ListIndex))
 lst_de.Visible = False: fr1.Enabled = True
  If dest = "14" Then
   MsgBox vbCrLf & "                *********  PLEASE NOTE  *********" & vbCrLf & vbCrLf & _
                   " IF Trailer is Destined for GUILBAULT, USE Code 80." & vbCrLf & vbCrLf, , "Manifest"
 End If
 If seal.Visible Then seal.SetFocus
End Sub
Private Sub lst_de_KeyDown(KeyCode%, Shf%)
 Select Case KeyCode
   Case vbKeyEscape
     lst_de.Visible = False: fr1.Enabled = True
     If dest.Visible Then dest.SetFocus
   Case vbKeyReturn, vbKeyExecute: lst_de_DblClick
 End Select
End Sub

Private Sub c_or_Click()
 Dim j As Integer
 If Val(orig) = 0 Then
   lst_or.ListIndex = -1
 Else
   For j = 0 To lst_or.ListCount - 1
     If Val(orig) = lst_or.ItemData(j) Then
       lst_or.ListIndex = j: Exit For
     End If
   Next j
 End If
 lst_or.Visible = True: fr1.Enabled = False: lst_or.SetFocus
End Sub
Private Sub lst_or_DblClick()
 orig = CStr(lst_or.ItemData(lst_or.ListIndex))
 lst_or.Visible = False: fr1.Enabled = True
 If dest.Visible Then dest.SetFocus
End Sub
Private Sub lst_or_KeyDown(KeyCode%, Shf%)
 Select Case KeyCode
   Case vbKeyEscape
     lst_or.Visible = False: fr1.Enabled = True
       If orig.Visible Then orig.SetFocus
   Case vbKeyReturn, vbKeyExecute: lst_or_DblClick
 End Select
End Sub

Private Sub l_edsav_Click()
 If ch_edsav Then ch_edsav = False Else ch_edsav = True
 If ch_edsav.Visible Then ch_edsav.SetFocus
End Sub
Private Sub ch_edsav_Click()
 If ch_edsav Then l_edsav.ForeColor = &H10 Else l_edsav.ForeColor = &H807080
End Sub
Private Sub l_prnt_Click(x%)
 o_prnt(x).Value = True
 If o_prnt(x).Visible Then o_prnt(x).SetFocus
End Sub
Private Sub l_prnt2_Click()
 o_prnt(1).Value = True
 If o_prnt(1).Visible Then o_prnt(1).SetFocus
End Sub
Private Sub o_prnt_Click(x%)
 Dim i As Integer
 For i = 0 To 2: l_prnt(i).ForeColor = &H706070: Next i
 l_prnt(x).ForeColor = &H20&
 If x = 0 Then l_prnt2.ForeColor = &H706070 Else l_prnt2.ForeColor = &H30&
 If x = 2 Then c_prnt.Caption = "Save Mnfst" Else c_prnt.Caption = "Print Mnfst"
End Sub

Private Sub ChkPrntr()
 Dim p As Printer
 On Error Resume Next
 Err = 0: lp = Printer.DeviceName 'look for printer name, whatever is default at start-up is maintained in program
 gateprn = "": defprn = ""
 If Err <> 0 Then 'no printers defined
   lp = "No Default Printer": c_prnt.Enabled = False: Err = 0: prnavail = False
 Else
   For Each p In Printers '21Nov13 look for dedicated Gate Pass printer if any
     Select Case p.DeviceName
       Case "GatePassPrinter"
         Set ogateprn = p: gateprn = "GatePassPrinter": prnavail = True
       Case "InlandInbondPrinter"
         Set oinbondprn = p: inbondprn = "InlandInbondPrinter"
       Case lp 'default
         Set odefprn = p: defprn = lp: prnavail = True
     End Select
   Next
 End If
End Sub

Private Function RunCmd(CmdPath As String, Optional WinStyle As VbAppWinStyle = vbNormalFocus) As Boolean
 Dim pc As Long
 'run external cmd script as synchronous process and wait for process to complete using API calls
 Dim hProc As Long   'external task ID#
 On Error GoTo RCERR 'trap all errors
 'run the script returning PID (usual Windows Task/Process ID no.)
 hProc = OpenProcess(synch, 0, Shell(CmdPath, WinStyle)) 'create a 'synch'ronous process so that VB is forced to wait for its completion
 If hProc Then 'if non-zero then script started, use WinAPI to wait for script to complete
   WaitForSingleObject hProc, infinit  'infinit = &HFFFF - max. possible wait time
   CloseHandle hProc 'well-behaved method to ensure process is fully killed
 Else 'process never started
   GoTo RCERR
 End If
 Exit Function
RCERR:
 Err.Clear
 RunCmd = True 'return true on any error
End Function

Private Function AuMail(mto$, sbj$) As Boolean
 Dim dm As String
 Dim fb As Object
 
 On Error GoTo PMERR
 
 'write mail body as text file
 If fo.FileExists(aempth & "body.txt") Then fo.DeleteFile aempth & "body.txt", True
 Set fb = fo.OpenTextFile(aempth & "body.txt", 8, True): fb.Write aembody: fb.Close
 'mail subject
 aemsubj = sbj
 'get recipients, make mail-to list
 aemto = "": rs.Open "select email from shpcntrptmailst where " & mto & "='1' order by email"
 If rs.EOF Then
   rs.Close: MsgBox "** FAIL Retreive Mail List for " & mto & " ! Aborting Mail Send . . .", , "Manifest": GoTo PMERR
 End If
 Do
   aemto = aemto & rs(0) & "; ": rs.MoveNext
 Loop Until rs.EOF
 rs.Close
 aemto = Left$(aemto, Len(aemto) - 2)
 
 If umail <> "" Then aemfrm = umail Else aemfrm = "david.young@speedy.ca"

 aemfrm = "david.young@speedy.ca"
 
 n = WritePrivateProfileString("SETUP", ByVal "From", ByVal aemfrm, aminipth)
 n = WritePrivateProfileString("SETUP", ByVal "To", ByVal aemto, aminipth)
 n = WritePrivateProfileString("SETUP", ByVal "Cc", ByVal "", aminipth)
 n = WritePrivateProfileString("SETUP", ByVal "Bcc", ByVal "", aminipth)
 n = WritePrivateProfileString("SETUP", ByVal "Subject", ByVal aemsubj, aminipth)
 n = WritePrivateProfileString("SETUP", ByVal "Charset", ByVal "windows-1252", aminipth)
 n = WritePrivateProfileString("SETUP", ByVal "TimeUnit", ByVal "S", aminipth)
 n = WritePrivateProfileString("SETUP", ByVal "Log", ByVal "0", aminipth)
 n = WritePrivateProfileString("SETUP", ByVal "Backup", ByVal "0", aminipth)
 n = WritePrivateProfileString("SETUP", ByVal "Delete", ByVal "1", aminipth)
 n = WritePrivateProfileString("SETUP", ByVal "PollInterval", ByVal "1", aminipth)
 n = WritePrivateProfileString("SETUP", ByVal "RunOnceExit", ByVal "1", aminipth)
 n = WritePrivateProfileString("SETUP", ByVal "AppendFileNames", ByVal "0", aminipth)
 n = WritePrivateProfileString("SETUP", ByVal "StartInTimerMode", ByVal "0", aminipth)
 n = WritePrivateProfileString("SETUP", ByVal "PollDir", ByVal aempoll, aminipth)

 Pause 0.25
 
 'run Auto Mail (synchronous, this prog suspended until user completes)
 If RunCmd(aempth & "automail.exe", vbNormalFocus) Then 'function returned Fail
   dm = "AutoMail Execute Failed!"
   MsgBox dm, , "Send Package"
 End If
 
 Pause 0.25
 
 'take auto-emailer out of run-once mode to facilitate manual execution/config
 n = WritePrivateProfileString("SETUP", ByVal "TimeUnit", ByVal "M", aminipth)
 n = WritePrivateProfileString("SETUP", ByVal "Delete", ByVal "0", aminipth)
 n = WritePrivateProfileString("SETUP", ByVal "PollInterval", ByVal "5", aminipth)
 n = WritePrivateProfileString("SETUP", ByVal "RunOnceExit", ByVal "0", aminipth)
 
 GoTo PMEND
 
PMERR:
  AuMail = True

PMEND:
 Set fb = Nothing
End Function

Private Sub c_prnt_Click()
 Dim dm As String, dm2 As String, sqs As String, BC As String, prolist As String
 Dim b As Boolean, embprnt As Boolean, dockpkup As Boolean, dlvd As Boolean
 Dim pndx As String, cd As String, retdte As String, delrec As String, towner As String
 Dim retrn As String, retftrun As String, retor As String, retde As String
 Dim rettpre As String, retupre As String, retfun As String, retftr As String, rtrip As String
 Dim j As Integer, totpros As Integer, a As Integer, c As Integer, k As Integer
 Dim totwt As Long
 Dim isloaded As Boolean, isdispatched As Boolean, sentinterline As Boolean, edoride As Boolean
 
 Dim tycop(1 To 104) As String        '22Aug12(DNY) Tyco Elect Markham auto-email
 Dim osm As String, dsm As String     '06Sep12(DNY) US Southbound closed trailer auto-email
 Dim sherway(1 To 104) As String      '15Oct14(DNY) - Sherway Group auto-email
 Dim schaeffsouth(1 To 104) As String '18Jul16(DNY) Schaeffler Can. to Schaeffler U.S. auto-email

'21Nov13 the following is derived from c_New_Click
' adjusted to compliment calls to c_prnt_Click

 gatepassprinted = False '^^
 gtesto = ""             '^^
 dockpkup = False
 If cnt = 0 Then 'no pros
  'fall-thru
 Else
   Select Case Val(orig)
     Case 31   'storage trailer - fall-thru
     Case 68, 37, 38, 69, 71, 72, 73 'terminal dock pickup/transfer codes -> printing done from c_New_Click
       dockpkup = True: savprnt = False
       MsgBox "SameDay, Dock Pickup/Transfer Printing, Auto-Deliver-Off, etc. are Handled when Saving Manifest", , "Manifest":  Exit Sub
     Case Else 'all others - fall-thru
   End Select
 End If
 
 savprnt = False
 If Not prnavail Then
   MsgBox "No Default Printer for this Computer!", , "Print Manifest": savprnt = False: Exit Sub
 End If

 If l_ne = "NEW" Then
    MsgBox "Manifest Must Be SAVED and Displayed in EDIT Mode to Print. Aborting Print . .", , "Manifest": savprnt = False: Exit Sub
 End If

 If Len(trip) <> 7 Then
   MsgBox "Invalid Trip No.! Please Contact IT Support. Aborting . .", , "Manifest": savprnt = False: Exit Sub
 End If

 '09Jan20(DNY) require Descrip entry for all Canadian LH/interline transfers
 If Val(orig) <> Val(dest) Then
   Select Case Val(orig)
     Case 1, 3, 4, 5, 6, 8, 10, 12, 25 To 28, 35, 85, 86, 91, 11
       Select Case Val(dest)  '
         Case 1, 3, 4, 5, 6, 8, 10, 12, 14, 25 To 28, 34, 35, 42, 43, 50, 51, 76, 77, 78, 79, 80, 81, 82, 85, 86, 87, 88, 91, 94, 95, 96, 11
           descr = Trim$(descr)
           If Len(descr) < 3 Then
             dm = "** 'Descrip' field Requires a minimum 3-Character Entry **  Aborting Save . . ."
             MsgBox dm, , "Description Check on Domestic Linehaul/Interline Schedules"
             savprnt = False: Exit Sub
           End If
       End Select
   End Select
 End If

 'count populated rows, get total weight, check for Tyco hi-vis mail trigger, 15Oct14 - check for Sherway Group standing apptmnt bills, send eMail with BOLs on 1-1 trip when trailer closed
 ' 21Jun19 - Home Hardware St. Jacobs, ON
 gpcnt = 0: totpros = 0: totwt = 0: a = 0: c = 0: Erase tycop: Erase sherway: Erase schaeffsouth: aembody = ""
 homehardware34henry = False
 For j = 0 To 206 Step 2
  If g2.TextMatrix(j, 17) <> "" And g2.TextMatrix(j, 3) <> "" Then
    totpros = totpros + 1: totwt = totwt + Val(g2.TextMatrix(j, 5))
    If mailsendon Then
      If Left$(g2.TextMatrix(j, 3), 18) = "TYCO ELECT MARKHAM" Then '22Aug12(DNY) TYCO ELECT MARKHAM auto-email trigger
        a = a + 1: tycop(a) = g2.TextMatrix(j, 17)
      End If
      If Val(orig) = 1 And Val(dest) = 1 And InStr(g2.TextMatrix(j, 3), "SHERWAY") > 0 And InStr(g2.TextMatrix(j, 3), "325 AN") > 0 And InStr(g2.TextMatrix(j, 3), "MISSISSAUG") > 0 Then
        c = c + 1: sherway(c) = g2.TextMatrix(j, 17) 'buff pros, build email body: pro - shipper shipper list
        aembody = aembody & g2.TextMatrix(j, 17) & " - " & g2.TextMatrix(j, 2) & vbCrLf
      End If
      If Val(orig) < 15 And Val(dest) > 49 And Left$(g2.TextMatrix(j, 2), 10) = "SCHAEFFLER" Then
        If Left$(g2.TextMatrix(j, 3), 10) = "SCHAEFFLER" Then
          c = c + 1: schaeffsouth(c) = g2.TextMatrix(j, 17)
        End If
      End If
      If Not homehardware34henry And Val(orig) = 1 And Val(dest) = 1 And InStr(g2.TextMatrix(j, 3), "HOME HARDW") > 0 And InStr(g2.TextMatrix(j, 3), "34 HENRY") > 0 And InStr(g2.TextMatrix(j, 3), "JACOB") > 0 Then homehardware34henry = True
    End If
  End If
 Next j
 twt = CStr(totwt): tsh = CStr(totpros)

SKPLHTRM:
 
'21Nov13 ----- 21Nov13 ----- 21Nov13 ----- 21Nov13 ----- 21Nov13 ----- 21Nov13 ----- 21Nov13 ----- 21Nov13 -----

' ****** >  >  If mef And o_prnt(2).Visible And o_prnt(2).Value Then GoTo SAVONLY

'21Nov13 ----- 21Nov13 ----- 21Nov13 ----- 21Nov13 ----- 21Nov13 ----- 21Nov13 ----- 21Nov13 ----- 21Nov13 -----

oprnt0 = False: oprnt1 = False: oprnt2 = False: oprntL = False: oprntG = False
gs1 = "": gs2 = "": gs3 = "": gs4 = "": gs5 = "": gs6 = "": gs7 = "": gs8 = "": gi = 0 'these are set from 'prntdisp' pop-up form handling
Printer.FillStyle = vbFSTransparent
 
isloaded = False: isdispatched = False: rtrip = ""
If rs.State <> 0 Then rs.Close
rs.Open "select cast(loaddatetime as char),cast(dispatchdatetime as char),loadedby,returntripnum from trip where tripnum='" & Right$(trip, 6) & "'"
If Not rs.EOF Then
  If IsDate(rs(0)) = True Then isloaded = True
  If IsDate(rs(1)) = True Then isdispatched = True
  rtrip = rs(3)
End If
rs.Close
 
 'auto-generate/update return trip from user prompt
'21Nov13 If o_prnt(1) Then
   Select Case Val(dest)
     Case 45, 46, 51 To 54, 55, 57 To 67
       Select Case Val(orig)
         Case 1, 3, 4, 5, 6, 8, 10, 12, 11
           '*** RULE: To set Return Trip, unit & trailer must be entered ***
           If Len(rtrip) = 6 And isdispatched Or isloaded Then
             'fall-thru
           Else
             If Not isloaded Then
               If Len(unit) = 0 Or Len(trailer) = 0 Then GoTo SKIPRETURNTRIP 'unit & trailer entry req'd
             End If
           End If
           gen1 = ""
           retrpfrm.Show 1 'prompt user for return trip - trip date (if any) returned in global var gen1
           On Error Resume Next
           If gen1 = "-999" And gen2 = "-999" Then
             MsgBox "Manifest Save/Print Aborted . . .", , "Trip Manifest": GoTo ENDCPRNT
           End If
           If gen1 = "" Then 'no return trip req'd
             If Val(rtrpn) > 0 Then 'return trip previously specified - remove return trip record/reference
               dm = "delete from trip where tripnum='" & rtrpn & "'": cmd.CommandText = dm: Err = 0: cmd.Execute
               If Err <> 0 Then
                 MsgBox "FAIL Remove Return Trip# " & rtrpn & "  Abort Save/Print . .", , "Manifest": Err = 0: GoTo ENDCPRNT
               End If
               dm = "update trip set returntripnum='' where tripnum='" & Right$(trip, 6) & "'"
               cmd.CommandText = dm: Err = 0: cmd.Execute
               If Err <> 0 Then
                 MsgBox "FAIL Remove Return Trip from Outbound Trip: " & Right$(trip, 6) & "  Abort Save/Print . .", , "Manifest": Err = 0: GoTo ENDCPRNT
               End If
               rtrpn = "" 'clear return trip value from screen buffer
             End If
           Else 'return trip specified by user
             If Val(rtrpn) = 0 Then 'no previous return trip record, create new
               rettrip = GetNxtTrpNum()
               If Val(rettrip) = 0 Then GoTo ENDCPRNT
               dm = "insert ignore into trip set tripnum='" & rettrip & "', ": dm2 = ""
             Else 'return trip already exists, force update to synch with outbound trip
               rettrip = rtrpn: dm = "update trip set": dm2 = " where tripnum='" & rettrip & "'"
             End If
             'save return trip - use same unit/trailer
             arg = FixUnitTrailerRun(retupre, retfun, rettpre, retftr, retrn, retftrun): retde = orig: retor = dest: ui$ = Left$(init, 1) & Right$(init, 1)
             hr$ = Format$(Now, "Hh"): min$ = Format$(Now, "Nn") 'hour/min last-modified
             sqs = " assign='" & Left$(trip, 1) & rettrip & "', date='" & gen1 & "', unit='" & retfun & _
                   "', trailer='" & retftr & "', run='" & retrn & "', seal='RETURN TRIP', dest_city='" & _
                    mySav(destcty(Val(retde))) & "', origin_city='" & mySav(destcty(Val(retor))) & _
                   "', distcode='" & ui$ & "', carrier='OUTBOUND: " & trip & "', origin='" & _
                    retor & "', dest='" & retde & "', date_modified='" & Format$(Now, "YYYY-MM-DD") & _
                   "', hr_mod='" & hr$ & "', min_mod='" & min$ & "', init='" & ui$ & _
                   "', unitpre='" & retupre & "', trlrpre='" & rettpre & "'"
             sqs = dm & sqs & dm2: cmd.CommandText = sqs: Err = 0: cmd.Execute
             If Err <> 0 Then
               MsgBox "FAIL Create/Update Return Trip!" & vbCrLf & Err.Description, , "Manifest": Err = 0
             End If
             'write return tripnum to outbound trip record
             sqs = "update trip set returntripnum='" & rettrip & "' where tripnum='" & Right$(trip, 6) & "'": cmd.CommandText = sqs: Err = 0: cmd.Execute: Err = 0
             dm2 = TrnTripHdr(rettrip) 'transfer trip header to Unix
             If Err <> 0 Then MsgBox Err.Description
             arg = TrnTripQC(rettrip, dm2) 'transfer trip QC record, display on screen
           End If
           On Error GoTo 0
       End Select
   End Select
'21Nov13  End If
SKIPRETURNTRIP:
 
 gpcnt = 0 'gatepass print count
 
 '** put any rules to skip the status pop-up selector here
 '   1rst part is here, 2nd part is after STBASEPRNT
 '1.a storage trailer
 If Val(orig) = 31 Then
   savprnt = True: c_new_Click
   If Not savprnt Then Exit Sub 'error during save, abort print
   GoTo STBASEPRNT
 End If
 '2.a sameday or pickup at dock - should not get here is dock pickup - print done from c_New_Click
 If dockpkup Then Exit Sub
 
 '3.a empty manifest
 If totpros = 0 Then
    dm = "Print Manifest/Gate Pass for this EMPTY Manifest ?"
    If MsgBox(dm, vbDefaultButton2 + vbQuestion + vbYesNo, "EMPTY Manifest Print") = vbNo Then Exit Sub
      
   '19May20(DNY) allow LH moves in USA between US Partner terminals
    Select Case Val(orig)
       Case 45, 46, 47, 51, 52, 53, 54, 55, 57 To 61, 63, 67  '19May2020 -> PYLE 45 to 47, EXLA 52 to 54, YRC 55, HMES 57 to 61, WARD 63, SEFL 67
          Select Case Val(dest)
             Case 45, 46, 47, 51, 52, 53, 54, 55, 57 To 61, 63, 67
                savprnt = True: c_new_Click
                cnewclick = False
                If Not savprnt Then Exit Sub 'save failed
                sqs = "update trip set dispatchdatetime='" & Format$(Now, "YYYY-MM-DD Hh:Nn:Ss") & "' where assign='" & trip & "'"
                cmd.CommandText = sqs: Err = 0: cmd.Execute
                GoTo STBASEPRNT
          End Select
    End Select
   
    'all other empty manifests
    savprnt = True
    If Val(orig) < 15 And Val(dest) < 15 Then cnewclick = False Else cnewclick = True
    c_new_Click
    cnewclick = False
    sqs = "update trip set dispatchdatetime='" & Format$(Now, "YYYY-MM-DD Hh:Nn:Ss") & "' where assign='" & trip & "'"
    cmd.CommandText = sqs: Err = 0: cmd.Execute
    If Err <> 0 Then
      MsgBox "FAIL Write Gate-Pass Dispatch Date-Time!" & vbCrLf & Err.Description, , "Manifest": Err = 0
    End If
    If Not savprnt Then Exit Sub 'error during save, abort print
   
    GoTo STBASEPRNT
 End If  '3.a
 
 '21Nov13 ----- 21Nov13 ----- 21Nov13 ----- 21Nov13 ----- 21Nov13 ----- 21Nov13 ----- 21Nov13 ----- 21Nov13 -----
 
PRNTSTRT:
 
 gs1 = "9": gi = totpros: rettrip = "": prngatpas = False
 prntdisp.Show 1
 If gs1 = "9" Then Exit Sub
 Select Case gs1 'set indicators to mimic old radio print option selection
   Case "1": oprnt0 = True 'print manifest, no dispatch
   Case "L": oprnt0 = True: oprntL = True 'print manifest, set Loading, no dispatch
   Case "G": oprnt1 = True: oprntG = True 'print GatePass, set On-Dlvy, send to Unix/dispatch
   Case "LG": oprnt1 = True: oprntL = True: prnt1 = True: oprntG = True
 End Select
 l_loadby = gs7 ' update loader on main screen
 
 '==========================================================' GoTo TESTSHERWAY  '===================================================================
 
 '18Sep2019(DNY) write report trigger for any linehaul transfer to PIC terminal
 '  updated 01Oct19(DNY) for PIC-XXX domestic linehauls
 '  updated 05Nov19(DNY) for MTL-XXX domestic linehauls
 '  updated 09Jan20(DNY) - All domestic linehauls and interline transfers
 
 sentinterline = False
 
 If InStr(gs1, "L") > 0 Or gs1 = "G" Then
      
   edoride = False
IPT0:

   If (Val(orig) <> Val(dest)) Then edoride = True
   If edoride = True Then
     Select Case Val(orig)
       Case 1, 3, 4, 5, 6, 8, 10, 11, 12, 25 To 28, 34, 35, 42, 43, 51, 76, 77, 78, 79, 80, 81, 82, 85, 86, 87, 88, 91, 94, 95, 96 'SZTG terminals + domestic partner interline codes
         Select Case Val(dest)  '
           Case 1, 3, 4, 5, 6, 8, 10, 11, 12, 25 To 28, 34, 35, 42, 43, 51, 76, 77, 78, 79, 80, 81, 82, 85, 86, 87, 88, 91, 94, 95, 96 'SZTG terminals + domestic partner interline codes
             If t_tro = "" Then towner = "SP" Else towner = t_tro
             dm = "insert ignore into closedtrailerreport set entrydt='" & Format$(Now, "YYYY-MM-DD Hh:Nn:Ss") & "',trip='" & gs6 & "',tripdate='" & Format$(mdte, "YYYY-MM-DD") & _
                  "',orig='" & orig & "',dest='" & dest & "',trailer='" & trailer & "',tractor='" & unit & "',loader='" & mySav(gs7) & "',init='" & init & _
                  "', trailerowner='" & towner & "', descrip='" & mySav(descr) & "', loadsht='" & mySav(ldsht) & "', seal='" & mySav(seal) & "', driver='" & mySav(driver) & "'"
             cmd.CommandText = dm: Err = 0: cmd.Execute
             If Err <> 0 Then
               MsgBox "FAIL Write Closed Trailer Update for Trip# " & tp & " Please Save this screen and send to IT ASAP." & vbCrLf & vbCrLf & Err.Description, , "Manifest": Err = 0
             End If
             If sentinterline Then GoTo SKPIPT1
         End Select
     End Select
   End If
   
 End If
ELHAML:
 If totpros = 0 Then GoTo STBASEPRNT
 
 '07Dec19(DNY) write email trigger on transfer manifests to interline partners
 If InStr(gs1, "G") > 0 Then
 
   Select Case Val(dest)
     Case 25 To 28, 34, 35, 77, 80, 85, 86, 87, 91 'domestic interline partners RSUT, KIDY, LAC/APOC, TGBT, AMOT-Monc, AMOT-Dart, DRU/DESS
       Select Case Val(orig)
         Case 1 To 12: GoTo IPT1  'ru   'TOR,PIC,LON, etc. -> RSUT WPG code 28  added 28May20(DNY), added RSUT CGY 27, EDM 26, VAN 25 codes for future support 02Jun20
       End Select
   End Select
   Select Case Val(orig)
     Case 25 To 28, 34, 35, 77, 80, 85, 86, 87, 91
       Select Case Val(dest)
          Case 1 To 12: GoTo IPT1
       End Select
   End Select
   GoTo SKPIPT1
   
IPT1:
   'create AMOT, RSUT interline transfer tracking record
   If t_tro = "" Then towner = "SP" Else towner = t_tro
   dm = "insert ignore into closedinterlinetrailers set entrydt='" & Format$(Now, "YYYY-MM-DD Hh:Nn:Ss") & "',trip='" & gs6 & "',tripdate='" & Format$(mdte, "YYYY-MM-DD") & _
        "',orig='" & orig & "',dest='" & dest & "',trailer='" & trailer & "',tractor='" & unit & "',loader='" & mySav(gs7) & "',init='" & init & _
        "', trailerowner='" & towner & "'"
   cmd.CommandText = dm: Err = 0: cmd.Execute
   If Err <> 0 Then
     MsgBox "FAIL Write Closed Interline Trailer Update for Trip# " & tp & _
            " Please Save this screen and send to IT ASAP." & vbCrLf & vbCrLf & Err.Description, , "Manifest": Err = 0
   End If
   If rs.State <> 0 Then rs.Close
   dm = "select recno from closedtrailerreport where trip='" & gs6 & "'"
   rs.Open dm
   If rs.EOF Then
      rs.Close: sentinterline = True: GoTo IPT0
   End If
   rs.Close
 End If
 
SKPIPT1:
 
 '21Jun19 trigger Home Hardware St.Jacobs mail alert on city trip 'Loading' status
 If homehardware34henry And oprntL Then
   SendMail Format$(Now, "DD-MMM-YYYY Hh:Nn") & "Hrs: Load Closed HOME HARDWARE ST. JACOBS, Trip " & Trim$(trip), "Trip " & Trim$(trip), "HomeHardwareStJacobsClosed"
 End If
 
 'save/update this manifest
 savprnt = True
 c_new_Click
 If Not savprnt Then Exit Sub 'error during save, abort print
 savprnt = False
 
STBASEPRNT:
 
 Disabl4Prn
 BC = Format$(mdte, "YYYYMMDD") & "SZTG" & Right$(trip, 6) 'create ACE style barcode for manifest & gate pass
 DrawBarCode BC, 0
 
 '1.b Storage - part 2 - print manifest (see 1.a above)
 If Val(orig) = 31 Then 'storage trailer
   pndx = GetPrnCtrl(): l_prnt_Click 1
   b = False: arg = PrntCpy(0, pndx, Format$(Now, "DDMMMYY-HhNn"), BC, "")
   GoTo ENDCPRNT
 End If
 '2.b Dock Pickup - should never get here, paperwork prints when saved - see c_New_Click
 '3.b empty manifest
 If totpros = 0 Then

   pndx = GetPrnCtrl(): l_prnt_Click 1
   b = False: arg = PrntCpy(0, pndx, Format$(Now, "DDMMMYY-HhNn"), BC, "")
   PrntGatePass 1, pndx, Format$(Now, "DDMMMYY-HhNn"), BC, 0
   
   Select Case Val(orig)  '25Mar20(DNY) send alert on empty southbound trailers for EXLA
      Case 1, 3, 4, 5, 6, 8, 10, 12 ', 11 only enable this when VAU replaces LAC for code 11
         Select Case Val(dest)
            Case 52 To 54
               rs.Open "select destcity from linehaul_status_codes where code='" & orig & "'"
               If Not rs.EOF Then osm = rs(0)
               rs.Close
               rs.Open "select destcity from linehaul_status_codes where code='" & dest & "'"
               If Not rs.EOF Then dsm = rs(0)
               rs.Close
               dm = vbCrLf & "Closed - EXLA Southbound Trailer: " & trailer & " (Empty)" & vbCrLf & vbCrLf & _
                             "Origin: " & osm & vbCrLf & vbCrLf & _
                             "  Dest: " & dsm & vbCrLf & vbCrLf
               If Len(trip) > 5 Then dm = dm & "  Trip: " & trip & vbCrLf & vbCrLf
               SendMail "** EXLA Southbound Closed Trailer: " & trailer & " (Empty)", dm, "EXLASouthManifested"
         End Select
   End Select
   
   
   GoTo ENDCPRNT
 End If

 Select Case gs1
   Case "1"    ' copyno,print ctrl#, date-time,       ,barcode strng, return trip#
     pndx = GetPrnCtrl(): l_prnt_Click 1
     For a = 1 To gi
       arg = PrntCpy(4, pndx, Format$(Now, "DDMMMYY-HhNn"), BC, "")
     Next a
     EnablPrn
     Exit Sub
     
   Case "L"
     pndx = GetPrnCtrl(): l_prnt_Click 1
     If Not PrntCpy(1, pndx, Format$(Now, "DDMMMYY-HhNn"), BC, rettrip) Then
       If Not PrntCpy(2, pndx, Format$(Now, "DDMMMYY-HhNn"), BC, rettrip) Then
         If Not PrntCpy(3, pndx, Format$(Now, "DDMMMYY-HhNn"), BC, rettrip) Then WrQStatus 'update trip record gs6=trip#, set/update Loaded 'Q' stathist record
       End If
     End If
     b = True: l_prnt_Click 0
         
   Case "G"  'gs3 = return trip no. if return trip gatepass req'd, gs4 = 1 = linehaul manifest, gi = #copies
     pndx = GetPrnCtrl(): l_prnt_Click 1
     For j = 1 To gi
       PrntGatePass 1, pndx, Format$(Now, "DDMMMYY-HhNn"), BC, CInt(gs4)
     Next j
     If Val(gs3) > 1000 Then 'return trip gate pass print;   test 768622 return from trip 768621 (BUF soutbound)
       rettrip = gs3
       BC$ = Format$(mdte, "YYYYMMDD") & "SZTG" & gs3
       DrawBarCode BC$, 0
       'gen2 = gate pass type; 1 = city, 2 = see prnttom, 3 = same-day return trip, 4 = next-day return
       ''''gen2 = "3" 'DEBUG DEBUG test same-day return trip gate pass print
       dm = orig: dm2 = dest: orig = dm2: dest = dm 'reverse orig/dest for return trip gate pass
       pndx = GetPrnCtrl()
       PrntGatePass Val(gen2), pndx, Format$(Now, "DDMMMYY-HhNn"), BC, "0"
       orig = dm: dest = dm2 'undo reversal
     End If
     b = True: l_prnt_Click 0
     WrOStatus '^^ update trip record gs6 = trip#, gs7=Loader, set/update out-on-dlvy 'O' or Linehaul 'L' status record
     
'' from prntdisp selection form
''   If islinehaul Then gs4 = "1" Else gs4 = "0"   'set gate pass type
''   If isinterline Then gs5 = "1" Else gs5 = "0"
     
   Case "LG"
     pndx = GetPrnCtrl(): l_prnt_Click 1
     If Not PrntCpy(1, pndx, Format$(Now, "DDMMMYY-HhNn"), BC, rettrip) Then
       If Not PrntCpy(2, pndx, Format$(Now, "DDMMMYY-HhNn"), BC, rettrip) Then
         If Not PrntCpy(3, pndx, Format$(Now, "DDMMMYY-HhNn"), BC, rettrip) Then WrQStatus 'update trip record gs6=trip#, set/update Loaded 'Q' stathist record
       End If
     End If
     For j = 1 To gi
       PrntGatePass 1, pndx, Format$(Now, "DDMMMYY-HhNn"), BC, CInt(gs4)
     Next j
     If Val(gs3) > 1000 Then 'return trip gate pass print;   test 768622 return from trip 768621 (BUF soutbound)
       rettrip = gs3:  BC$ = Format$(mdte, "YYYYMMDD") & "SZTG" & gs3
       DrawBarCode BC$, 0
       dm = orig: dm2 = dest: orig = dm2: dest = dm 'reverse orig/dest for return trip gate pass
       pndx = GetPrnCtrl(): PrntGatePass Val(gen2), pndx, Format$(Now, "DDMMMYY-HhNn"), BC, "0"
       orig = dm: dest = dm2 'undo reversal
     End If
     b = True: l_prnt_Click 0
     WrOStatus '^^update trip record gs6 = trip#, gs7=Loader, set/update out-on-dlvy 'O' or 'L'inehaul status record
     
 End Select
 
 '^^
 'only check city trips
 Select Case Val(orig)
   Case 1, 3, 4, 5, 6, 8, 10, 12, 11
     If Val(dest) <> Val(orig) Then GoTo SKPGPPCHK
   Case Else: GoTo SKPGPPCHK
 End Select
 If gtesto = "" Then
   For j = 0 To 206 Step 2
     If g2.TextMatrix(j, 17) <> "" And g2.TextMatrix(j, 3) <> "" Then
       gtesto = g2.TextMatrix(j, 17): Exit For
     End If
   Next j
   If gtesto = "" Then GoTo SKPGPPCHK 'empty manifest, skip alerts
 End If
 If Len(mnfst!trip) = 7 Then gs6 = Right$(mnfst!trip, 6) Else gs6 = mnfst!trip
 If gtesto <> "" And gatepassprinted Then
   sqs = "select pro from stathist where pro='" & gtesto & "' and date_status='" & Format$(mdte, "YYYY-MM-DD") & "' and statuscode='O'"
   If rs.State <> 0 Then rs.Close
   rs.Open sqs
   If rs.EOF Then
     rs.Close
     WrOStatus
   Else
     rs.Close
   End If
 Else
   If gatepassprinted Then WrOStatus
 End If
 dm3 = vbCrLf & "Trip:" & trip & "  Pro:" & gtesto & "  Init:" & init & "  IP:" & wsck.LocalIP & vbCrLf
 If gtesto <> "" And gatepassprinted Then
   sqs = "select pro from stathist where pro='" & gtesto & "' and date_status='" & Format$(mdte, "YYYY-MM-DD") & "' and statuscode='O'"
   If rs.State <> 0 Then rs.Close
   rs.Open sqs
   If rs.EOF Then
      SendMail "GatePass No-Save Alert3", CStr(dm3), "GatePassSaveAlert" 'send e-mail via SMTP - vars: subject line, primary message, mySQL mail-to fieldname, pronumber
      dm = "** IMPORTANT **  -  Gate Pass Reccord is Not Complete, On-Delivery Record Will NOT TRANSFER to Envoy" & vbCrLf & vbCrLf & _
           "              ** To FIX This - Simply RE-PRINT the Gate Pass Until this Message No Longer Appears  **" & vbCrLf & vbCrLf & _
           " Note: an alert e-mail has automatically been sent to IT with the trip details."
      MsgBox dm
   End If
   rs.Close
 Else
   If gatepassprinted Then
      SendMail "GatePass No-Save Alert + WrOStatus By-Passed", CStr(dm3), "GatePassSaveAlert" 'send e-mail via SMTP - vars: subject line, primary message, mySQL mail-to fieldname, pronumber
      dm = "** IMPORTANT **  -  Gate Pass Reccord is Not Complete, On-Delivery Record Will NOT TRANSFER to Envoy" & vbCrLf & vbCrLf & _
           "              ** To FIX This - Simply RE-PRINT the Gate Pass Until this Message No Longer Appears  **" & vbCrLf & vbCrLf & _
           " Note: an alert e-mail has automatically been sent to IT with the trip details."
      MsgBox dm
   End If
 End If
SKPGPPCHK:
 gatepassprinted = False: gtesto = ""
'^^
 
' 21Nov13 ----- 21Nov13 ----- 21Nov13 ----- 21Nov13 ----- 21Nov13 ----- 21Nov13 ----- 21Nov13 ----- 21Nov13 ----- 21Nov13 -----
' moved 21Nov13(DNY) to accomodate LOADED/ON-DLVY separate GatePass print features
'from 22Aug12(DNY) Tyco Elect Markham auto-email
 If oprntG = True And Val(orig) = 1 And Val(dest) = 1 And Len(tycop(1)) > 5 Then
   dm = vbCrLf & "Trailer Closed destined to Tyco Markham. Please ensure a picture of the" & vbCrLf & _
        "trailer is taken prior to departing the Brampton yard." & vbCrLf & vbCrLf & _
        "Trailer No. " & t_tro & trailer & vbCrLf & "Pro(s) on load - "
   For a = 1 To 104
     If tycop(a) = "" Then Exit For
     dm = dm & tycop(a) & ", "
   Next a
   If Right$(dm, 2) = ", " Then dm = Left$(dm, Len(dm) - 2)
   dm = dm & vbCrLf & vbCrLf & "Please reply to all cc'd with the pictures and ensure product is secured" & _
             vbCrLf & "for delivery."
   SendMail "**HI-VIS  TYCO MARKHAM - Trip " & trip, dm, "TycoElectMarkham"
 End If
 
 'from 06Sep12(DNY) Close any US Southbound trailer
 If oprntG Then
   Select Case Val(dest)
      Case 45, 46
        rs.Open "select destcity from linehaul_status_codes where code='" & orig & "'"
        If Not rs.EOF Then osm = rs(0)
        rs.Close
        rs.Open "select destcity from linehaul_status_codes where code='" & dest & "'"
        If Not rs.EOF Then dsm = rs(0)
        rs.Close
        prolist = ""
        dm = vbCrLf & "Closed - PYLE Southbound Trailer: " & trailer & vbCrLf & vbCrLf & _
                      "Origin: " & osm & vbCrLf & vbCrLf & _
                      "  Dest: " & dsm & vbCrLf & vbCrLf
        If Len(trip) > 5 Then dm = dm & "  Trip: " & trip & vbCrLf & vbCrLf
        For j = 0 To 204 Step 2
          If Len(g2.TextMatrix(j, 17)) > 5 And g2.TextMatrix(j, 3) <> "" Then prolist = prolist & "  " & g2.TextMatrix(j, 17) & vbCrLf
        Next j
        If prolist <> "" Then dm = dm & prolist
        SendMail "** PYLE Southbound Closed Trailer: " & trailer & " on Trip: " & trip, dm, "PYLESouthManifested"
     Case 52 To 54
       rs.Open "select destcity from linehaul_status_codes where code='" & orig & "'"
       If Not rs.EOF Then osm = rs(0)
       rs.Close
       rs.Open "select destcity from linehaul_status_codes where code='" & dest & "'"
       If Not rs.EOF Then dsm = rs(0)
       rs.Close
       dm = vbCrLf & "Closed - EXLA Southbound Trailer: " & trailer & vbCrLf & vbCrLf & _
                     "Origin: " & osm & vbCrLf & vbCrLf & _
                     "  Dest: " & dsm & vbCrLf & vbCrLf
       If Len(trip) > 5 Then dm = dm & "  Trip: " & trip & vbCrLf & vbCrLf
       SendMail "** EXLA Southbound Closed Trailer: " & trailer, dm, "EXLASouthManifested"
     Case 88 'Hartrans - email e-manifest in JSON format, track in edimail990tender under scac=HART, pro=Trip#
       If Val(orig) < 12 Then
'         If SendJSONManifest(88) Then
'           'arg=arg
'         End If
       End If
     Case 1, 8
       If (Val(orig) = 1 Or Val(orig) = 8) And Left$(trailer, 2) = "TH" Then
         dm = "Please Confirm Closing Interline Transfer to Tom Hartnett Peterborough."
         If MsgBox(dm, vbDefaultButton1 + vbExclamation + vbYesNo, "e-Manifest Trigger for Hartrans") = vbYes Then
'         If SendJSONManifest(88) Then
'           'arg=arg
'         End If
         End If
       End If
   End Select
 End If
 If oprntG And Val(orig) > 0 And (Val(dest) = 45 Or Val(dest) = 46 Or (Val(dest) > 54 And Val(dest) < 68)) Then
   rs.Open "select destcity from linehaul_status_codes where code='" & orig & "'"
   If Not rs.EOF Then osm = rs(0)
   rs.Close
   rs.Open "select destcity from linehaul_status_codes where code='" & dest & "'"
   If Not rs.EOF Then dsm = rs(0)
   rs.Close
   dm = vbCrLf & "Closed - US Southbound Trailer: " & trailer & vbCrLf & vbCrLf & _
                 "Origin: " & osm & vbCrLf & vbCrLf & _
                 "  Dest: " & dsm
   SendMail "** US Southbound Closed Trailer: " & trailer, dm, "ussouthmanifest"
   
 '18Jul16(DNY) Schaeffler-Schaeffler southbound auto-email
   If schaeffsouth(1) <> "" Then
     dm = vbCrLf & "Pro(s) " & schaeffsouth(1) & vbCrLf
     For c = 2 To 104 'get all Schaeffler pros
       If schaeffsouth(c) <> "" Then dm = dm & "       " & schaeffsouth(c) & vbCrLf
     Next c
     dm = dm & vbCrLf & "***** Please Ensure the Bond is Saved into Protrace *****" & vbCrLf & vbCrLf
     SendMail "**Hi-Vis SCHAEFFLER-to-SCHAEFFLER Southbound Bond on Trailer " & trailer, dm, "SchaefflerBond"
   End If
 End If
 
 '15Oct14 trigger Sherway Milton standing apptmnt eMail on set city trip 'Loading' status
 If c > 0 And sherway(1) <> "" And oprntL Then
   If fo.FolderExists("c:\temp\stpc\automail") Then fo.DeleteFolder "c:\temp\stpc\automail", True: Pause 0.25
   fo.CreateFolder "c:\temp\stpc\automail": Pause 0.25
   For c = 1 To 104 'cycle thru manifest & put copies of Sherway BOL scans to email attachment folder
     If sherway(c) = "" Then Exit For
     For a = 1 To 19 'pull up to 19 pages of BOL
       dm = "select ndx from img_stpc where pro_date='" & sherway(c) & "' and inv_unit='BOL' and pgno='" & CStr(a) & _
            "' and accnt <> 'VOID' order by scandatetime desc limit 1"
       If ir.State <> 0 Then ir.Close
       ir.Open dm
       If ir.EOF Then
         ir.Close: Exit For
       End If
       dm5$ = ir(0): dm4$ = "c:\temp\stpc\automail\" & sherway(c) & "-" & Format$(a, "000") & ".tif" 'buff image index, set filename
       ir.Close
       arg = MGetTempBlobImage(dm5$, sherway(c), "2", dm4$) 'index, pro, type, filename
     Next a
   Next c
   aembody = "Attached BOL - Shipper:" & vbCrLf & aembody
   dm6$ = Format$(Now, "DD-MMM-YYYY Hh:Nn") & "Hrs: Load Closed SHERWAY WHSE 8PM Standing Apptmnt"
   arg = AuMail("sherwaycc", dm6$) 'mailto fieldname, subject text
 End If
 
ENDCPRNT:
 frw.Visible = False
 MousePointer = 0
 fr1.Enabled = True: frg1.Enabled = True
 c_new.Enabled = True: c_prnt.Enabled = True
 c_clr.Enabled = True: c_fgo.Enabled = True
 If Not embprnt Then c_prnt.SetFocus 'not embedded manifest print
 
SAVONLY:

End Sub
         
Private Function SendJSONManifest(p&) As Boolean
 Dim dm As String, b As String
 Dim cnt As Integer
 Dim f As Object
 
 Select Case p
   Case 88 ' Hartrans
     For j = 0 To 206 Step 2
       If Len(g2.TextMatrix(j, 17)) > 6 Then
         b = b & "," & vbCrLf 'if *not* the last shipment, put comma after shipment section close ->  {,
         dm = BuildJSONShipment(88, j)
         If dm = "EOF" Then
           SendJSONManifest = True: Exit Function
         Else
           b = b & dm
         End If
       End If
     Next j
     If Len(b) > 10 Then
       b = b & "   ]" & vbCrLf & "}" & vbCrLf
       b = BuildJSONManifestHeader(88) & b
       Set f = fo.OpenTextFile("c:\json-test.txt", 8, True)
       f.Write b
       f.Close: Set f = Nothing
     End If
   Case Else
 End Select
End Function

Private Function BuildJSONShipment$(p&, j)
 Dim dm As String, h As String, q As String, pcstype As String
 Dim i As Integer, pcs As Integer, scnt As Integer
 Dim s() As String
 
 q = Chr(34)
 If rsc.State <> 0 Then rsc.Close
 Select Case p
   Case 88
     '                     0            1            2        3          4           5      6     7     8     9       10       11
     '                12                   13        14      15
     rsc.Open "select cons_company,cons_address,cons_city,cons_prov,cons_postcode,cartons,skids,drums,pails,other,other_desc,totalwt," & _
              "cast(apptmnt_date as char),door,apptmnt_time,remark from probill where pronumber='" & g2.TextMatrix(j, 17) & "'"
     If rsc.EOF Then
       BuildJSONShipment = "EOF"
     Else
       h = h & "         " & q & "SZTG Pronumber" & q & ":" & q & g2.TextMatrix(j, 17) & q & "," & vbCrLf
       h = h & "         " & q & "Consignee Name" & q & ":" & q & rsc(0) & q & "," & vbCrLf
       h = h & "         " & q & "Consignee Address" & q & ":" & q & rsc(0) & q & "," & vbCrLf
       h = h & "         " & q & "Consignee City" & q & ":" & q & rsc(0) & q & "," & vbCrLf
       h = h & "         " & q & "Consignee Province" & q & ":" & q & rsc(0) & q & "," & vbCrLf
       h = h & "         " & q & "Consignee Postal" & q & ":" & q & rsc(0) & q & "," & vbCrLf
       For i = 5 To 9
         If Val(rsc(i)) > 0 Then
           pcs = Val(rsc(i))
           Select Case i
             Case 6: pcstyp = "PAL": Exit For
             Case 5: pcstyp = "CTN": Exit For
             Case 7: pcstyp = "DRM": Exit For
             Case 8: pcstyp = "RLL": Exit For
             Case 9: pcstyp = "PCS": Exit For
           End Select
         End If
       Next i
       h = h & "         " & q & "Pcs" & q & ":" & q & CStr(pcs) & q & "," & vbCrLf
       h = h & "         " & q & "Pcs Units" & q & ":" & q & pcstype & q & "," & vbCrLf
       h = h & "         " & q & "Tot Wt" & q & ":" & q & rsc(11) & q & "," & vbCrLf
       h = h & "         " & q & "Wt Units" & q & ":" & q & "lbs" & q & "," & vbCrLf
       If IsDate(rsc(12)) = True Then
         h = h & "         " & q & "Delivery Appt Date (YYYY-MM-DD)" & q & ":" & q & rsc(12) & q & "," & vbCrLf
         Select Case rsc(13)
           Case 1: dm = "ASAP"
           Case 0 To 2400: dm = Format$(Val(rsc(13)), "0000") & " Hrs ET"
           Case Else: dm = ""
         End Select
         If dm <> "" Then h = h & "         " & q & "Delivery Appt Time/Zone 1" & q & ":" & q & dm & q & "," & vbCrLf
         Select Case Val(rsc(15))
           Case 1 To 2400: dm = rsc(15)
           Case Else: dm = ""
         End Select
         If dm <> "" Then h = h & "         " & q & "Delivery Appt Time/Zone 2" & q & ":" & q & dm & q & "," & vbCrLf
       End If
       scnt = 0
       If Len(g2.TextMatrix(j, 7)) > 4 Then 'rcvng hours
         scnt = scnt + 1
         h = h & "         " & q & "Special Requirements " & CStr(scnt + 1) & q & ":" & q & g2.TextMatrix(j, 7) & q & "," & vbCrLf
       End If
       If Len(g2.TextMatrix(j + 1, 4)) > 5 And InStr(g2.TextMatrix(j + 1, 4), "**") > 0 Then 'special instructions, eg. ** TAILGATE REQ'D ** STRAIGHT TRUCK REQ'D **
         s = Split(g2.TextMatrix(j + 1, 4), "**")
         For i = 0 To UBound(s)
           scnt = scnt + 1
           h = h & "         " & q & "Special Requirements " & CStr(scnt + 1) & q & ":" & q & Trim$(s(i)) & q & "," & vbCrLf
         Next i
       End If
       n = n & "      {" & vbCrLf
     End If
     rsc.Close
   Case Else
 End Select
 BuildJSONShipment = h
End Function
         
Private Function BuildJSONManifestHeader$(p&)
 Dim h As String, q As String
 q = Chr(34)
 h = "{" & vbCrLf
 Select Case p
   Case 88 'Hartrans
     h = h & "   " & q & "Sender Company" & q & ":" & q & "Speedy Transport" & q & "," & vbCrLf
     h = h & "   " & q & "Sender SCAC" & q & ":" & q & "SZTG" & q & "," & vbCrLf
     h = h & "   " & q & "Document Type" & q & ":" & q & "Interline Manifest" & q & "," & vbCrLf
     h = h & "   " & q & "Manifest Date (YYYY-MM-DD)" & q & ":" & q & Format$(mdte, "YYYY-MM-DD") & q & "," & vbCrLf
     h = h & "   " & q & "Manifest Reference Number" & q & ":" & q & Right$(trip, 6) & q & "," & vbCrLf
     h = h & "   " & q & "Shipment" & q & ":[" & vbCrLf
     h = h & "      " & q & "{" & vbCrLf
   Case Else
 End Select
 BuildJSONManifestHeader = h
End Function

Private Sub WrQStatus()
 Dim dm As String, dm2 As String, dm3 As String, mi As String, b As Boolean
 On Error Resume Next
 dm = "update trip set loaddatetime='" & Format$(Now, "YYYY-MM-DD Hh:Nn:Ss") & "', loadinit='" & init & "', loadedby='" & mySav(gs7) & _
      "' where tripnum='" & gs6 & "'"
 cmd.CommandText = dm: Err = 0: cmd.Execute
 If Err <> 0 Then
   MsgBox "FAIL Trip Loading Update Trip# " & tp & vbCrLf & vbCrLf & Err.Description, , "Manifest": Err = 0: b = True
 End If
 If Minute(Now) = 0 Then mi = "0" Else mi = CStr(Minute(Now) - 1)
 For j = 0 To 206 Step 2
   If g2.TextMatrix(j, 17) <> "" And g2.TextMatrix(j, 3) <> "" Then
     rs.Open "select recno from stathist where pro='" & g2.TextMatrix(j, 17) & "' and date_status='" & Format$(mdte, "YYYY-MM-DD") & _
            "' and statuscode='Q' and origin='" & orig & "' and dest='" & dest & "' and trailer='" & trailer & _
            "' order by date_ent desc, cast(hr_ent as unsigned) desc, cast(min_ent as unsigned) desc limit 1"
     dm2 = "date_status='" & Format$(mdte, "YYYY-MM-DD") & "', statuscode='Q', origin='" & orig & "', dest='" & dest & "', deltimepalm='" & init & _
           "', hr_ent='" & Format$(Hour(Now), "00") & "', min_ent='" & mi & "', date_ent='" & Format$(Now, "YYYY-MM-DD") & _
           "', trailer='" & trailer & "', ship_company='" & mySav(gs2) & "', cons_city='" & gs6 & "'"
     If rs.EOF Then
       dm = "insert ignore into stathist set pro='" & g2.TextMatrix(j, 17) & "', " & dm2
     Else
       dm = "update stathist set " & dm2 & " where recno='" & rs(0) & "'"
     End If
     rs.Close
     cmd.CommandText = dm: Err = 0: cmd.Execute
     If Err <> 0 Then
       MsgBox "FAIL Set Q-Loading Status for " & g2.TextMatrix(j, 17) & vbCrLf & vbCrLf & Err.Description, , "Manifest": Err = 0
     End If
   End If
 Next j
 
End Sub

Private Sub WrOStatus() 'update trip with dispatch, loader info, set 'O'ut-on-Dlvy status
 Dim dm As String, dm2 As String, dm3 As String, stcod As String, mi As String

 On Error Resume Next
 
 ''09Jun20(DNY)  If gatepassprinted And gs7 = "" And Val(orig) = Val(dest) Then GoTo SKPTRPUPD '^^
 
 dm = "update trip set dispatchdatetime='" & Format$(Now, "YYYY-MM-DD Hh:Nn:Ss") & "', dispatchinit='" & init & "', loadedby='" & mySav(gs7) & _
      "' where tripnum='" & gs6 & "' "
 cmd.CommandText = dm: Err = 0: cmd.Execute
 If Err <> 0 Then
   MsgBox "FAIL Trip Dispatch Update Trip# " & tp & vbCrLf & vbCrLf & Err.Description, , "Manifest": Err = 0
 End If

 If Val(orig) = 67 Then    'stop any on-delivery status write for a Charlotte-direct SEFL shipment
   Select Case Val(dest)
     Case 1, 3, 4, 5, 6, 8, 10, 11, 12: Exit Sub
   End Select
 End If
 
SKPTRPUPD:

 If Minute(Now) + 1 > 59 Then mi = "59" Else mi = CStr(Minute(Now) + 1)
 For j = 0 To 206 Step 2
  
   '//If gs4 = "1" Or gs5 = "1" Then stcod = "L" Else stcod = "O" 'deprecated 19Jan21(DNY)
   stcod = "L"
   Select Case CInt(orig)
     Case 1, 3, 4, 5, 6, 8, 10, 11, 12, 25 To 28, 34, 85, 86
       Select Case CInt(dest)
         Case 1, 3, 4, 5, 6, 8, 10, 11, 12, 25 To 28, 34, 85, 86
           If orig = dest Then stcod = "O"
       End Select
   End Select
   If stcod = "L" And GetTermDisp(CInt(orig)) <> "" And GetTermDisp(CInt(dest)) <> "" Then
     gs8 = "IN-TRANSIT " & GetTermDisp(CInt(dest)) & " ex-" & GetTermDisp(CInt(orig)) 'matches std. Evoship In-Transit linehaul status message text
   End If
   
   If g2.TextMatrix(j, 17) <> "" And g2.TextMatrix(j, 3) <> "" Then
   
     If gtesto = "" Then gtesto = g2.TextMatrix(j, 17) '^^
     
     rs.Open "select recno from stathist where pro='" & g2.TextMatrix(j, 17) & "' and date_status='" & Format$(mdte, "YYYY-MM-DD") & _
             "' and statuscode='" & stcod & "' order by date_ent desc, cast(hr_ent as unsigned) desc, cast(min_ent as unsigned) desc limit 1"
     dm2 = "date_status='" & Format$(mdte, "YYYY-MM-DD") & "', statuscode='" & stcod & "', origin='" & orig & "', dest='" & dest & _
           "', deltimepalm='" & init & "', hr_ent='" & Format$(Hour(Now), "00") & "', min_ent='" & mi & _
           "', date_ent='" & Format$(Now, "YYYY-MM-DD") & "', trailer='" & trailer & "', ship_company='" & mySav(gs8) & "', cons_city='" & gs6 & "'"
     If rs.EOF Then
       dm = "insert ignore into stathist set pro='" & g2.TextMatrix(j, 17) & "', " & dm2
     Else
       dm = "update stathist set " & dm2 & " where recno='" & rs(0) & "'"
     End If
     rs.Close: cmd.CommandText = dm: Err = 0: cmd.Execute
     If Err <> 0 Then
       MsgBox "FAIL Set " & stcod & " Status for " & g2.TextMatrix(j, 17) & vbCrLf & vbCrLf & Err.Description, , "Manifest": Err = 0
     End If
   End If
 Next j
End Sub

Private Function SetCorresp(p$, nit$, cat$, tx$) As Boolean
 Dim aconn As New ADODB.Connection
 Dim ars As New ADODB.Recordset
 Dim acmd As New ADODB.Command
 Dim dm As String, ndx As String
 Dim ndxv As Long
 
 On Error Resume Next
 aconn.Open aconnstr: ars.ActiveConnection = aconn: acmd.ActiveConnection = aconn: acmd.CommandType = adCmdText

 srcinfo = wsck.LocalIP & " " & mySav(wsck.LocalHostName)
 acmd.CommandText = "update ndx_counter set ndx=last_insert_id(ndx+1)": Err = 0: acmd.Execute
 ars.Open "select last_insert_id()"
 If Err = 0 Then ndxv = ars(0) 'buffer image index no.
 If ars.State <> 0 Then ars.Close
 If Err <> 0 Then
   MsgBox "INTERNAL ERROR: " & Err.Description, , "ProTrace": Err = 0: Exit Function
 End If
 ndx = Format$(ndxv, "000000000") 'build path index string from index no.
 dm = "insert into filemaster set ndx='" & ndxv & "', pro='" & p & "', filetype='Comment', adddate='" & _
       Format$(Now, "YYYY-MM-DD") & "', addtime='" & Format$(Now, "Hh:Nn:Ss") & "', comment='" & mySav(tx) & _
      "', category='" & cat & "', init='" & nit & "', access='1', srcinfo='" & srcinfo & _
      "', lastevent='AddCom', delblock='0'"
 acmd.CommandText = dm: Err = 0: acmd.Execute
 If Err <> 0 Then
   LS = Format$(Now, "DD-MMMYY Hh:Nn") & " FAIL Correspondence " & Err.Description: Err = 0: SetCorresp = True
 End If
 Set ars = Nothing: Set acmd = Nothing: aconn.Close: Set aconn = Nothing
 
End Function

Private Function PrntCpy(cpy%, cn$, cd$, BC$, rettrip$) As Boolean 'main print routine - copyno, print control no., date-time, return trip#
 Dim curr As Integer, str As Integer, enr As Integer, pcs As Integer
 Dim i As Integer, j As Integer, k As Integer, n As Integer
 Dim pg As Integer, totpg As Integer, std As Integer, totpros As Integer, lhdl As Integer
 Dim sk As Integer, ct As Integer, ot As Integer 'specific piece counts
 Dim dm As String, dm2 As String, ms As String, seg1 As String, seg2 As String, trs As String
 Dim segs() As String
 Dim pcseg() As String
 Dim b As Boolean, quikx As Boolean, clrke As Boolean, fast As Boolean, meyt As Boolean, ward As Boolean
        ''27Sep16  tlog As Boolean,

 totpros = Val(tsh)
 On Error Resume Next
 
 'get total print pages
 If totpros = 0 Then
   totpg = 1
 Else
   totpg = Fix(totpros / 12)
   If totpg * 12 < totpros Then totpg = totpg + 1
 End If
 pg = 1: i = 0 'init pg#, starting row to process
 ward = (Val(dest) = 63)  '''meyt = (Val(dest) = 89):
 Printer.Duplex = 0      'single-sided printing
 Printer.Orientation = 2 'force landscape mode
 Printer.ScaleMode = 5   'scale to inches,  set base print font
 Printer.FontTransparent = True: Printer.DrawWidth = 3
 
 'determine whether delivery, linehaul or storage manifest
 If Val(orig) = Val(dest) Then
   If Val(orig) = 31 Then lhdl = 2 Else lhdl = 0
 Else
   Select Case Val(dest)
     Case 71, 72, 73: lhdl = 0
     Case Else: lhdl = 1
   End Select
 End If

REMFPAGE:

 Printer.Orientation = 2 'force landscape mode
 Printer.ScaleMode = 5   'scale to inches,  set base print font
 Printer.FontTransparent = True: Printer.DrawWidth = 3
    
 'ACE type barcode center-top
 Printer.PaintPicture barc(0).Image, 11# / 2# - 1.5, 0.27, 3#, 0.31 'x1, y1, Width, Height
 Printer.FontSize = 9
 cprnt 11# / 2#, 0.58, BC
  
 'print manifest sheet form background
 Printer.FontSize = 10
 PrntMnfstBkGrnd lhdl, cpy
 
 'write page no. right-top
 Printer.FontSize = 10
 dm = "Page " & CStr(pg) & " of/de " & CStr(totpg)
 lprnt (9.5 + 0.9) - Printer.TextWidth(dm), 0.55, dm
 'page no. center-top
 Printer.FontName = "Arial Black"
 Printer.FontSize = 14
 Printer.ForeColor = &H808080
 cprnt 11# / 2!, 0, dm
 Printer.FontName = "Arial"
 Printer.FontSize = 9
 Printer.ForeColor = &H0
 dm = UCase$(cn & "-" & Format$(Now, "DDMMMYY-HhNn") & "-" & init)
 lprnt 7.25, 0.58, dm
 
 Printer.FontSize = 10
 'write total wt. / cube wt.
 dm = twt
 If Val(tcu) > 0 Then dm = dm & " (" & tcu & " cu.)"
 cprnt 6.0625 + ((7.8125 - 6.0625) / 2!), 7# + 0.375, dm
 'write date
 Printer.FontSize = 12
 cprnt 1.4375 / 2!, 0.92, Format$(mdte, "DD-MMM-YYYY")
 Printer.FontSize = 10
 'write driver name if entered
 If Trim$(driver) <> "" Then
   Printer.FontSize = 8
   cprnt 1.4375 + ((3! - 1.4375) / 2!), 0.97, Trim$(driver)
   Printer.FontSize = 10
 End If
 'write unit
 dm = unit
 If Trim$(t_uno) <> "" Then
   dm = t_uno & dm
 Else
   If InStr(t_un, " ST ") > 0 Then
     dm = "ST " & dm
   ElseIf InStr(t_un, " CUBE ") > 0 Then
     dm = "CVAN " & dm
   End If
 End If
 cprnt 3! + ((4.125 - 3!) / 2!), 0.95, dm
 'write trailer
 trs = trailer
 If Trim$(t_tro) <> "" Then trs = t_tro & trs
 dm2 = ""
 If Val(trlrapp) > 1 Then
   Select Case Val(trlrapp)
     Case 2: dm2 = "2nd"
     Case 3: dm2 = "3rd"
     Case Else: dm2 = trlrapp & "th"
   End Select
   dm = trs & " - " & dm2  ''& " Run" takes up too much space
 Else
   dm = trs
 End If
 cprnt 4.125 + ((5.5625 - 4.125) / 2!), 0.82, dm
 'write seal if entered
 If Trim$(seal) <> "" Then cprnt 4.125 + ((5.5625 - 4.125) / 2!), 0.95, Trim$(seal)
 '21Nov13 write loader
 Printer.FontSize = 9: dm = Replace(Trim$(l_loadby), "  ", " ")
 If InStr(dm, " ") Then
   Do
     If Printer.TextWidth(dm) < 1.6875 Then Exit Do
     Printer.FontSize = Printer.FontSize - 1
   Loop Until Printer.FontSize < 7
   cprnt 5.5625 + ((7.25 - 5.5625) / 2!), 0.95, dm
    ''' cprnt 5.5625 + ((7.25 - 5.5625) / 2!), 0.89, segs(0): cprnt 5.5625 + ((7.25 - 5.5625) / 2!), 0.99, segs(1)
 Else
   
   If Printer.TextWidth(dm) > 1.6875 Then Printer.FontSize = 8
   If Printer.TextWidth(dm) > 1.6875 Then Printer.FontSize = 7
   If Printer.TextWidth(dm) > 1.6875 Then Printer.FontSize = 6
   cprnt 5.5625 + ((7.25 - 5.5625) / 2!), 0.95, dm
 End If
 'write trip#
 Printer.FontSize = 10
 If Len(trip) < 7 Then dm = "- UNASSIGNED -" Else dm = trip
 dm = dm & "  " & l_or
 If l_de = l_or Then dm = dm & " Local" Else dm = dm & " to " & l_de
 dm = dm & "  Tot-" & tsh: cprnt 7.25 + ((10.375 - 7.25) / 2!), 0.95, "Trip " & dm '9.5
 
 If rettrip <> "" Then
   If Replace(Trim$(l_loadby), "  ", " ") = "" Then '21Nov13 if not loader, write return trip there else move to left of barcode
     dm = "Return Trip: " & Left$(trip, 1) & rettrip: cprnt 5.5625 + ((7.25 - 5.5625) / 2!), 0.95, dm
   Else
     lprnt 7.3, 0.15, "Return Trip:":  lprnt 7.3, 0.33, Left$(trip, 1) & rettrip
   End If
 End If
 
 'write datalines
 k = -1  'init datarow count for this page
 voff = 0.3126 + 0.15
 For j = i To 206 Step 2  'cycle thru trip pros - successive pages start at 'i'
   
   If g2.TextMatrix(j, 17) = "" Then GoTo NXTPRNTJ

   '04Apr14(DNY) - if posted 'All-Short' - then auto un-post & set correspondence that located in this trailer
   If oprntL Or oprntG Then 'on statuses: loading/loaded/transferring or dispatching
     rs.Open "select recno from procnl where pro='" & g2.TextMatrix(j, 17) & "' and endinit=''" 'check if posted all-short
     If Not rs.EOF Then 'found, un-post
       seg2 = trailer
       If Trim$(t_tro) <> "" Then seg2 = t_tro & seg2
       seg1 = "update procnl set enddt='" & Format$(Now, "YYYY-MM-DD Hh:Nn:Ss") & "', endinit='" & init & "', endterm='" & uterm & "', " & _
             "location='Trlr " & trs & "' where pro='" & g2.TextMatrix(j, 17) & "'": cmd.CommandText = seg1: cmd.Execute
       arg = SetCorresp(g2.TextMatrix(j, 17), init, "OS&D", "Product Located: " & uterm & " Trailer " & trs & " Trip: " & Right$(trip, 6))
     End If
     rs.Close: Pause 0.01
   End If

   '                     0           1             2           3        4           5          6      7     8     9    10       11          12
   rs.Open "select ship_company,cons_company,cons_address,cons_city,cons_prov,cons_postcode,cartons,skids,drums,pails,other,cons_accnum,ship_accnum" & _
           " from probill where pronumber='" & g2.TextMatrix(j, 17) & "'"
   If Not rs.EOF Then
     
     k = k + 1 'increment row#
     
     'pronumber - 21Apr17 EXLA print width base is 0.9", keep reducing font size until pro# fits
     Do While Printer.TextWidth(g2.TextMatrix(j, 17)) > 0.9
       Printer.FontSize = Printer.FontSize - 0.25
     Loop
      ''lprnt 0.875 - Printer.TextWidth(g2.TextMatrix(j, 17)), 1.5625 + (k * voff) + 0.01, g2.TextMatrix(j, 17)
     lprnt 0.89 - Printer.TextWidth(g2.TextMatrix(j, 17)), 1.5625 + (k * voff) + 0.01, g2.TextMatrix(j, 17)
     Printer.FontSize = 10
     
     
     '' 27Sep16  If Not quikx And Not tlog And Not fast And Not meyt And Not ward Then 'leave space under pro for delivery agent partner beyond pro
     If Not quikx And Not fast And Not meyt And Not ward Then 'leave space under pro for delivery agent partner beyond pro
       If g2.TextMatrix(j, 12) = "1" Then 'early delivery
         Printer.FontSize = 8: Printer.FontItalic = True
         lprnt 0.875 - Printer.TextWidth("Early Delivery"), 1.5625 + (k * voff) - 0.12, "Early Delivery"
         Printer.FontSize = 10: Printer.FontItalic = False
       End If
     End If
     'shipper
     Printer.FontSize = 8.5
     seg1 = Trim$(rs(0)): seg2 = ""
     If Printer.TextWidth(seg1) > 1.25 Then TSplit 1.25, seg1, seg2
     If seg2 <> "" Then
       lprnt 0.9375 + 0.0625, 1.5625 - 0.08 + (k * voff), seg1 '0.07
       lprnt 0.9375 + 0.0625, 1.5625 + 0.03 + (k * voff), seg2 '0.04
     Else
       lprnt 0.9375 + 0.0625, 1.5625 + 0.02 + (k * voff), seg1 '0.02
     End If
     'consignee
     seg1 = Trim$(rs(1)): seg2 = ""
     If Printer.TextWidth(seg1) > 1.25 Then TSplit 1.25, seg1, seg2
     If seg2 <> "" Then
       lprnt 2.25 + 0.0625, 1.5625 - 0.08 + (k * voff), seg1
       lprnt 2.25 + 0.0625, 1.5625 + 0.03 + (k * voff), seg2
     Else
       lprnt 2.25 + 0.0625, 1.5625 + 0.02 + (k * voff), seg1
     End If
     'address
     seg1 = Trim$(rs(2)): seg2 = ""
     If Printer.TextWidth(seg1) > 1.25 Then TSplit 1.25, seg1, seg2
     If seg2 <> "" Then
       lprnt 3.5625 + 0.0625, 1.5625 - 0.08 + (k * voff), seg1
       lprnt 3.5625 + 0.0625, 1.5625 + 0.03 + (k * voff), seg2
     Else
       lprnt 3.5625 + 0.0625, 1.5625 + 0.02 + (k * voff), seg1
     End If
     'city
     seg1 = Trim$(rs(3)): seg2 = ""
     If Printer.TextWidth(seg1) > 1.0625 Then TSplit 1.0625, seg1, seg2
     If seg2 <> "" Then
       lprnt 4.875 + 0.0625, 1.5625 - 0.08 + (k * voff), seg1
       lprnt 4.875 + 0.0625, 1.5625 + 0.03 + (k * voff), seg2
     Else
       lprnt 4.875 + 0.0625, 1.5625 + 0.02 + (k * voff), seg1
     End If
     'prov
     lprnt 6.125 + 0.0625, 1.5625 + 0.02 + (k * voff), rs(4)
     Printer.FontSize = 10
     'COD or Collect
     Printer.FontSize = 9
     If g2.TextMatrix(j, 8) <> "" Then ''19=codcharges, 20=collect,24=billedparty
       n = InStr(g2.TextMatrix(j, 8), Chr(13))
       If n > 0 Then
         cprnt 6.78125, 1.5625 - 0.08 + (k * voff), Left$(g2.TextMatrix(j, 8), n - 1)
         cprnt 6.78125, 1.5625 + 0.03 + (k * voff), Mid$(g2.TextMatrix(j, 8), n + 1)
       End If
     End If
     'pcs counts for any of 3 types - cartons, skids, other (drums, pails, other)
     pcseg = Split(g2.TextMatrix(j, 4), Chr(13))
     n = UBound(pcseg)
     ct = 0: ot = 0: sk = 0
     For n = 0 To UBound(pcseg)
       If pcseg(n) <> "" Then
         arg = Left(pcseg(n), InStr(pcseg(n), "-") - 1)
         Select Case Right$(pcseg(n), 1)
           Case "C": ct = arg
           Case "S": sk = arg
           Case "O": ot = arg
         End Select
       End If
     Next n
     If ct > 0 And sk > 0 And ot > 0 Then 'at least 1 of each type
       Printer.FontSize = 8
       lprnt 7.5625 - 0.0625 - Printer.TextWidth(CStr(ct)), 1.5625 - 0.12 + (k * voff), CStr(ct)
       lprnt 7.5625 - 0.0625 - Printer.TextWidth(CStr(ot)), 1.5625 - 0.02 + (k * voff), CStr(ot)
       lprnt 7.5625 - 0.0625 - Printer.TextWidth(CStr(sk)), 1.5625 + 0.08 + (k * voff), CStr(sk)
       lprnt 7.5625 + 0.03125, 1.5625 - 0.12 + (k * voff), "C"
       lprnt 7.5625 + 0.03125, 1.5625 - 0.02 + (k * voff), "O"
       lprnt 7.5625 + 0.03125, 1.5625 + 0.08 + (k * voff), "S"
     ElseIf (ct > 0 And sk > 0) Or (ct > 0 And ot > 0) Or (sk > 0 And ot > 0) Then 'any 2 types
       Printer.FontSize = 8.5
       If ct > 0 Then
         dm = CStr(ct): dm2 = "C"
         If ot > 0 Then
           dm3$ = CStr(ot): dm4$ = "O"
         Else
           dm3$ = CStr(sk): dm4$ = "S"
         End If
       Else
         dm = CStr(ot): dm2 = "O"
         dm3$ = CStr(sk): dm4$ = "S"
       End If
       lprnt 7.5625 - 0.0625 - Printer.TextWidth(dm), 1.5625 - 0.08 + (k * voff), dm
       lprnt 7.5625 - 0.0625 - Printer.TextWidth(dm3$), 1.5625 + 0.04 + (k * voff), dm3$
       lprnt 7.5625 + 0.03125, 1.5625 - 0.08 + (k * voff), dm2
       lprnt 7.5625 + 0.03125, 1.5625 + 0.04 + (k * voff), dm4$
     ElseIf ct > 0 Or sk > 0 Or ot > 0 Then 'only 1 type
       If ct > 0 Then
         dm = CStr(ct): dm2 = "C"
       ElseIf ot > 0 Then
         dm = CStr(ot): dm2 = "O"
       Else
         dm = CStr(sk): dm2 = "S"
       End If
       Printer.FontSize = 8.5
       lprnt 7.5625 - 0.0625 - Printer.TextWidth(dm), 1.5625 + (k * voff), dm
       lprnt 7.5625 + 0.03125, 1.5625 + (k * voff), dm2
     Else 'none
       dm = "0": lprnt 7.5625 - 0.0625 - Printer.TextWidth(dm), 1.5625 + (k * voff), dm
     End If
     
     Printer.FontSize = 10
     dm = Replace(g2.TextMatrix(j, 5), Chr(13), "")
     lprnt 8.25 - 0.0625 - 0.0625 - Printer.TextWidth(dm), 1.5625 + (k * voff), dm 'wt
     
     'cube wt from hidden cell
     dm = Trim$(g2.TextMatrix(j, 14))
     If Val(dm) > 0 Then cprnt 8.5, 1.5625 + (k * voff), dm   'wt
          
     'delivery agent partner pro#
     If quikx Then ''27Sep16 Or tlog Then '****************************
       Printer.FontSize = 9
       Select Case Val(Left$(g2.TextMatrix(j, 20), 1))
         Case 5 To 9
           If quikx Then dm = "QXTI " '' Else dm = "NDTL "
         Case Else
           If quikx Then dm = "QTIM " '' Else dm = "NDTL "
       End Select
       dm = dm & g2.TextMatrix(j, 20)
       lprnt 0.875 - Printer.TextWidth(dm), 1.5625 + (k * voff) + 0.17, dm
       Printer.FontSize = 10
     End If
     If fast Then
       Printer.FontSize = 8
       dm = "FAST " & g2.TextMatrix(j, 20)
       lprnt 0.91 - Printer.TextWidth(dm), 1.5625 + (k * voff) + 0.17, dm
       Printer.FontSize = 10
     End If
     If meyt Then
       Printer.FontSize = 8
       dm = "MEYT " & g2.TextMatrix(j, 20)
       lprnt 0.91 - Printer.TextWidth(dm), 1.5625 + (k * voff) + 0.17, dm
       Printer.FontSize = 10
     End If
     If ward Then
       Printer.FontSize = 8
       dm = "WARD " & g2.TextMatrix(j, 20)
       lprnt 0.91 - Printer.TextWidth(dm), 1.5625 + (k * voff) + 0.17, dm
       Printer.FontSize = 10
     End If
     
     'last location {@}
     Printer.FontName = "Arial Narrow": Printer.FontSize = 9
     dm = Replace(Replace(Trim$(g2.TextMatrix(j + 1, 0)), "?", "-"), ":", "")
     If Len(dm) > 23 Then
       n = InStr(dm, "-T-")
       If n = 0 Then n = InStr(dm, "-R-")
       n = InStr(n + 3, dm, " ")
       dm = Left$(dm, n - 1)
     End If
     lprnt 0.9575, 1.5625 + 0.19 + (k * voff), dm
     Printer.FontName = "Arial": Printer.FontSize = 10
     
     'accnt codes
     If Trim$(g2.TextMatrix(j + 1, 3)) <> "" Then '{@}
       ''27Sep16 If Not quikx And Not tlog And Not fast And Not meyt And Not ward Then
       If Not quikx And Not fast And Not meyt And Not ward Then
         Printer.FontSize = 14
         If InStr(g2.TextMatrix(j + 1, 3), "HAZMAT") > 0 Then  '{@}
           lprnt 0.1, 1.5625 + (k * voff) + 0.17, "**"
           Printer.FontSize = 8: Printer.FontItalic = True
           lprnt 0.3, 1.5625 + (k * voff) + 0.19, "HAZMAT"
         Else
           cprnt 0.9375 / 2!, 1.5625 + (k * voff) + 0.17, "**"
         End If
       End If
       Printer.FontSize = 10: Printer.FontItalic = True
       cprnt (10# - 2.25) / 2# + 2.25, 1.5625 + 0.18 + (k * voff), Trim$(g2.TextMatrix(j + 1, 3)) '{@}
       Printer.FontItalic = False
     End If
     
     'check for auto-email coding on consignee
     '21Nov13 If o_prnt(0).Value = False And cpy = 1 Then 'if not a trial print (ie. manifest will be sent to dispatch)
     If oprnt1 Then 'if dispatching = printing Gate Pass . . .
       Select Case Val(orig) 'check for auto-email directives
         Case 1, 3, 4, 5, 6, 8, 10, 12 '11 future VAU 'currently only city runs - may have to open up later
           If Val(dest) = Val(orig) Then
             rsb.Open "select msgascons1 from scops where accno='" & rs(11) & "' and msgcons1asremark='1' and msgrmrkmnfstmail='1'"
             If Not rsb.EOF Then
               dm = "** Hi-Vis " & rs(1) & " Shipment Manifested for " & Format$(mdte, "DDD DD-MMM-YYYY")
               ms = vbCrLf & "Hi-Vis Delivery Instruction for Pro " & g2.TextMatrix(j, 17) & _
                    ":" & vbCrLf & vbCrLf & rsb(0) & vbCrLf & vbCrLf & "Consignee:" & vbCrLf & _
                    rs(1) & vbCrLf & rs(2) & vbCrLf & rs(3) & "  " & rs(4) & "  " & rs(5)
               SendMail dm, ms, "m" & rs(11) & "dlvyinstr" 'SMTP protocol - pass subject, message body, mySQL fieldname for mail recipients
             End If
             rsb.Close
           End If
       End Select
     End If
     
     'appointment
     If g2.TextMatrix(j, 6) <> "" Then
       If g2.TextMatrix(j, 12) = "1" Then 'marked for early delivery, gray bkgrnd slighly inside existing lines
         Printer.Line (8.8125 + 0.015, 1.4375 + 0.004 + (k * 0.4625))-(9.5 - 0.015, 1.4375 - 0.003 + (k * 0.4625) + 0.3125), &HC0C0C0, BF
       End If
       Printer.FontSize = 8.5
       segs = Split(g2.TextMatrix(j, 6), Chr(13))
       n = UBound(segs)
       Select Case n
         Case 0 'single line entry
           cprnt 9.15625, 1.5625 - 0.02 + (k * voff), g2.TextMatrix(j, 6)
         Case 1, 2
           cprnt 9.15625, 1.5625 - 0.13 + (k * voff), Format$(DateValue(segs(0)), "DD-MMM-YY")
           cprnt 9.15625, 1.5625 - 0.02 + (k * voff), segs(1)
           If n = 2 And segs(2) <> "" Then cprnt 9.15625, 1.5625 + 0.07 + (k * voff), segs(2)
       End Select
     End If
     
     Printer.FontSize = 10
     'Rcvg Hrs
     If g2.TextMatrix(j, 7) <> "" Then
       n = InStr(g2.TextMatrix(j, 7), "-")
       If n > 0 Then
         cprnt 10.2, 1.5625 - 0.11 + (k * voff), Replace(Left$(g2.TextMatrix(j, 7), n - 1), Chr(13), "") '08
         cprnt 10.2, 1.5625 + 0.04 + (k * voff), Mid$(g2.TextMatrix(j, 7), n + 1) '06
       Else
         cprnt 10.2, 1.5625 - 0.1 + (k * voff), g2.TextMatrix(j, 7) '01
       End If
     End If
   End If
   rs.Close
   
   If k = 11 Then '12 datarows per page max
     If totpros > (pg * (k + 1)) Then
       pg = pg + 1: Printer.EndDoc: i = j + 2: GoTo REMFPAGE 'increment page, print last page, set start row
     End If
     Exit For
   Else 'manifest had less than 17 datarows
     std = 0 'indicate to finish printing
   End If
   
NXTPRNTJ:

 Next j
  
EDOC:
  
 Printer.EndDoc 'output current page
 
 MousePointer = 0
End Function

Private Sub DrawBarCode(str$, x%)
 Dim objBC As clsBarCode39    'invoke barcode print-to-picturebox class object
 Set objBC = New clsBarCode39 'call the class into memory
 With objBC 'pass properties (vertical/left margins default 0 (VOffset, LOffset)
   .ShowLabel = False    'don't display value underneath barcode (disabled in class anyways)
   .ShowStartStop = True 'print start/top bars to improve reading
   .LineWeight = 1
   .TextString = str     'pass string as barcode value
   .Refresh barc(x)      'print barcode to picturebox with class code
 End With
 Set objBC = Nothing
End Sub

Private Sub TSplit(wid!, t1$, t2$)
 Dim tstr() As String 'text segment array
 Dim dm As String
 Dim n As Integer, p As Integer
 'split text into 2 segments to fit column width as necessary
 tstr = Split(t1, " ")
 dm = ""
 For n = 0 To UBound(tstr)
   dm = dm & " " & tstr(n)
   If Printer.TextWidth(Trim$(dm)) > wid Then
     dm = ""
     For p = 0 To n - 1
       dm = dm & tstr(p) & " "
     Next p
     Exit For
   End If
 Next n
 If n > UBound(tstr) Then GoTo JMPSPLT
 t1 = Trim$(dm)
 dm = ""
 For p = n To UBound(tstr)
   dm = dm & tstr(p) & " "
 Next p
 t2 = Trim$(dm)
 Do While Printer.TextWidth(t2) > wid
  t2 = Left$(t2, Len(t2) - 1)
 Loop
 Exit Sub
JMPSPLT:
 t2 = ""
End Sub

Private Sub PrntGatePass(typ%, cn$, cd$, BC$, lh%)
 'typ = trip type where 3 = same-day return trip, 4 = next-day return trip
 'cn = print control no. (next-number placed on gate-pass matching the parent manifest print-out)
 'cd = print date
 'BC = barcode string/value
 'lh = linehaul manifest=1, other=0
 Dim i As Integer, j As Integer, procnt As Integer
 Dim apptcnt As Integer, apptfirst As Integer, apptlast As Integer, js As Single
 Dim oy As Single
 Dim dm As String, dm2 As String, upre As String, tpre As String, tplate As String
 Dim unp As String, trp As String, stcity As String, encity As String
 Dim apptfirsttime As String, apptlasttime As String
 Dim appta() As String
 Dim specreq As Boolean
 
 If gateprn = "GatePassPrinter" Then
   Set Printer = ogateprn: SetDefaultPrinter "GatePassPrinter"
 ElseIf inbondprn = "InlandInbondPrinter" Then
   Set Printer = oinbondprn: SetDefaultPrinter "InlandInbondPrinter"
 End If
 
 Printer.FontTransparent = True
 Printer.Orientation = 1 'force protrait mode
 Printer.ScaleMode = 5 'inches
  
 'ACE type barcode center-mid
 Printer.PaintPicture barc(0).Image, 1.5, 3.875, 5!, 1! 'x1, y1, Width, Height
 Printer.FontName = "Consolas": Printer.FontSize = 10
 cprnt 8# / 2#, 3.875 + 1.06 - 0.05, BC
 
 Printer.PaintPicture pic1, 0#, 0.05, 2.05, 0.6154
 Printer.DrawWidth = 12
 Printer.Line (0.01, 0.75)-(8#, 3.75), , B 'main box
 Printer.DrawWidth = 3
 Printer.Line (5.25, 0.75)-Step(0, 3#) 'vert
 For j = 1 To 5
   Printer.Line (0.01, 0.75 + (j * 0.5))-Step(5.25, 0#)
   If j = 1 Then Printer.Line (5.25, 0.75 + (j * 0.5))-Step(2.75, 0#)
   If j = 5 Then Printer.Line (5.25, 0.75 + (j * 0.5))-Step(2.75, 0#) ')(
 Next j
 Printer.Line (5.25, 1.75)-Step(2.75, 0#)
 Printer.Line (5.25, 2.75)-Step(2.75, 0#)
 
 Printer.FontName = "Arial": Printer.FontBold = True: Printer.FontSize = 18
 lprnt 3.5, 0.42, "GATE PASS"
 Printer.FontSize = 16
 Select Case typ
   Case 3 To 4 'same/next-day return trip
     cprnt 5.25 + (2.75 / 2#), 0.25, Left$(trip, 1) & rettrip
   Case Else
     cprnt 5.25 + (2.75 / 2#), 0.25, trip
 End Select
 Printer.FontSize = 12
 lprnt 0.2, 0.75 + 0.17, "TRUCK CO."
 lprnt 0.2, 1.25 + 0.17, "TRUCK No."
 lprnt 0.2, 1.75 + 0.17, "TRAILER CO."
 lprnt 0.2, 2.25 + 0.17, "TRAILER No."
 lprnt 0.2, 2.75 + 0.17, "DRIVERS SIGNATURE"
 lprnt 0.2, 3.25 + 0.17, "AUTHORIZED BY"
 
 Printer.FontSize = 11
 cprnt 5.25 + (2.75 / 2#), 0.75, "DATE"
 cprnt 5.25 + (2.75 / 2#), 1.25, "LOADER"
 cprnt 5.25 + (2.75 / 2#), 1.75, "ADDITIONAL INFO"
 cprnt 5.25 + (2.75 / 2#), 2.75, "SEAL NUMBER"
 cprnt 5.25 + (2.75 / 2#), 3.25, "TRAILER PLATE"
 
 Printer.FontSize = 18
 'unit
 unp = unit
 If IsNumeric(unit) = True Then
   Select Case Trim$(t_uno)
     Case "", "SP", "SDT"  'sdt*
       If InStr(t_un, " ST ") > 0 Then
         unp = "ST " & unp
       ElseIf InStr(t_un, " CUBE ") > 0 Then
         unp = "CVAN " & dm
       End If
     Case Else: unp = t_uno & unp
   End Select
 End If
 lprnt 2.75, 1.25 + 0.11, unp
  
 'trailer
 trp = trailer: dm2 = ""
 Select Case Trim$(t_tro)
   Case "", "SP", "SDT", "HOL", "NPM", "ODF", "YRC", "YRT"  'sdt*
   Case Else: trp = t_tro & trp
 End Select
 If Val(trlrapp) > 1 Then
   Select Case Val(trlrapp)
     Case 2: dm2 = "2nd"
     Case 3: dm2 = "3rd"
     Case Else: dm2 = trlrapp & "th"
   End Select
   trp = trp & " - " & dm2 & " Run"
 End If
 lprnt 2.75, 2.25 + 0.11, trp
 
 'attempt to discern unit/trailer company
 upre = "": tpre = ""
 Select Case Trim$(t_uno)
   Case "", "SP", "SDT": upre = "Speedy Transport"  'sdt*
   Case Else
     For j = 0 To 99
       If uownr(j, 0) = "" Then Exit For
       If uownr(j, 0) = t_uno Then
         upre = uownr(j, 1): Exit For
       End If
     Next j
 End Select
 
 Select Case Trim$(t_tro)
   Case "", "SP", "SDT" 'sdt*
     tpre = "Speedy Transport"
   Case Else
      For j = 0 To 99
       If tsc(j, 0) = "" Then Exit For
       If tsc(j, 0) = t_tro Then
         tpre = tsc(j, 1): Exit For
       End If
     Next j
 End Select

 Printer.FontSize = 16
 If unit <> "" And upre <> "" Then lprnt 2.75, 0.75 + 0.11, upre
 If tpre <> "" Then lprnt 2.75, 1.75 + 0.11, tpre
 
 'seal
 Select Case typ
   Case 3 To 4
     cprnt 5.25 + (2.75 / 2#), 3.15, "Return of " & trip
   Case Else
    '' cprnt 5.25 + (2.75 / 2#), 3.15, seal
     cprnt 5.25 + (2.75 / 2#), 3.15 - 0.1875, seal ')(
 End Select

 'authorized by: gate pass issuer 07Apr16(DNY)
 lprnt 1.75, 3.37, gpissuer
 
 If Trim$(mnfst!l_loadby) <> "" Then
   Printer.FontSize = 13: js = 0.22
   Do
     Printer.FontSize = Printer.FontSize - 1: js = js + 0.01
     If Printer.TextWidth(l_loadby) <= 2.73 Then Exit Do
   Loop Until Printer.FontSize = 9
   cprnt 5.25 + (2.75 / 2#), 1.25 + js, l_loadby
   Printer.FontSize = 16
 End If
 'date
 Select Case typ
   Case 4 '4=next-day return trip
     ''21Nov13 cprnt 5.25 + (2.75 / 2#), 1.15, Format$(mdte + 1, "DD-MMM-YYYY")
     cprnt 5.25 + (2.75 / 2#), 0.94, Format$(mdte + 1, "DD-MMM-YYYY")
   Case Else '3=same-day return trip
     '' 21Nov13 cprnt 5.25 + (2.75 / 2#), 1.15, Format$(mdte, "DD-MMM-YYYY")
     cprnt 5.25 + (2.75 / 2#), 0.94, Format$(mdte, "DD-MMM-YYYY")
 End Select
 
 Printer.FontSize = 9
 cprnt 5.25 + (2.75 / 2#), 0.55, UCase$(cn & "-" & cd & "-" & init) 'print control no.

'''''''GoTo SKPLOCL '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

 'special print processing for local trips only
 stcity = "": encity = "": procnt = 0: apptcnt = 0: apptfirst = 0: specreq = False: signer = ""
 Select Case Val(orig)
   Case 1, 3, 4, 5, 6, 8, 10, 12  '11 future VAU
     If Val(dest) = Val(orig) Then
       'gather info: first-shipment city, end-shipment city, appts, first appt time/shipment no.
       For j = 0 To 206 Step 2
         If Val(g2.TextMatrix(j, 17)) > 0 Then 'shipment entry in this row
           procnt = procnt + 1
           If Not specreq And Trim$(g2.TextMatrix(j + 1, 3)) <> "" Then specreq = True
           If g2.TextMatrix(j, 6) <> "" Then 'apptmnt
             apptcnt = apptcnt + 1 'increment appt count
           End If
           'get consignee city for shipment
           appta = Split(g2.TextMatrix(j, 3), Chr(13))
           If UBound(appta) > 1 Then
             i = InStr(appta(2), "  ")
             If i > 0 Then dm = Trim$(Left$(appta(2), i)) Else dm = Trim$(appta(2))
           Else
             dm = "Not Entered"
           End If
           If stcity = "" Then stcity = dm Else encity = dm
         End If
       Next j
       If procnt = 1 Then encity = stcity
       'print info
       Printer.FontSize = 11
       lprnt 5.35, 2.1, "Tot Shipments - " & tsh: lprnt 5.62, 2.3, "Tot Weight - " & twt
       If apptcnt = 0 Then
         dm = "None"
       Else
         dm = CStr(apptcnt)
         If apptcnt = Val(tsh) Then dm = CStr(apptcnt) & " (All)"
       End If
       lprnt 5.699, 2.5, "Apptmnts - " & dm
       Printer.FontSize = 13: lprnt 1.1 - 0.5, 5.23, "Start:": lprnt 4.5 - 0.25, 5.23, "End:"
       Printer.FontSize = 24: lprnt 1.65 - 0.5, 5.1, stcity: lprnt 5.01 - 0.25, 5.1, encity
       Printer.FontSize = 12: Printer.FontItalic = True
       If apptcnt > 0 Then lprnt 1.1 - 0.85, 5.5, "* Appointments on Board - see Manifest"
       If specreq Then lprnt 1.1 - 0.85, 5.75, "*** Special Requirements - see Manifest"
       Printer.FontName = "Arial Black": Printer.FontSize = 27 '27Jul16(DNY)
       cprnt 4!, 9.6, "* Deploy ENVOY Before Leaving Yard *"
       Printer.FontName = "Arial": Printer.FontSize = 13: Printer.FontItalic = False
     End If
 End Select
 
 If Val(orig) < 20 And Val(orig) = Val(dest) Then
   trailerloc.Refresh
   pic2.Picture = LoadPicture("\\132.147.119.126\ProTrace\TrailerLastLocationImage\000000_NotFound.jpg")
   dm = Trim$(trailer): i = Len(dm)
   If i > 4 Then
     For j = 0 To trailerloc.ListCount - 1
       If Left$(trailerloc.List(j), i + 1) = dm & "_" Then
         pic2.Picture = LoadPicture("\\132.147.119.126\ProTrace\TrailerLastLocationImage\" & trailerloc.List(j))
         dm = Mid$(trailerloc.List(j), i + 2, 14)
         If Len(dm) > 11 Then
           dm = "As of " & Format$(Mid$(dm, 5, 2), "MMM") & "-" & Mid$(dm, 7, 2) & " " & Mid$(dm, 9, 2) & ":" & Mid$(dm, 11, 2) & "Hrs"
           Printer.FontSize = 10: Printer.FontBold = True: Printer.FontItalic = False
           rprnt 3.65, 7.5, dm
         End If
         Exit For
       End If
     Next j
   End If
   Printer.PaintPicture pic2, 3.75, 5.5, 4!, 4!
   GoTo EPGP
 End If
 
 Select Case Val(orig)
   Case 1, 3, 4, 5, 6, 8, 10, 12, 14, 51, 56  'add 11 when VAU opened, replaces LAC
   Case Else: GoTo EPGP
 End Select
  Select Case Val(dest)
   Case 1, 3, 4, 5, 6, 8, 10, 12, 14, 35, 50, 51, 56, 77, 78, 79, 80, 87, 91 'add 11 when VAU opened, replaces LAC
   Case Else: GoTo EPGP
 End Select
 
  
 'print layout lines (uses Speedy DR format)
 Printer.DrawWidth = 3
 PrntGPProBkGrnd -1, 5.2 'bottom copy
  
 Printer.FontSize = 10: Printer.FontBold = True: Printer.FontItalic = True
  
 'ship location
 '             0    1    2    3      4     5
 dm = "select nam,addr,city,prov,postcode,tel from linehaul_status_codes where code='" & orig & "'"
 rs.Open dm
 If rs.EOF Then
   rs.Close: MsgBox "FAIL Get Location for Origin Code: " & orig & ". Cannot Print Linehaul DR!", , "Manifest": GoTo EPGP
 End If
 'cons location
 dm2 = "select nam,addr,city,prov,postcode,tel from linehaul_status_codes where code='" & dest & "'"
 rsb.Open dm2
 If rs.EOF Then
   rsb.Close: MsgBox "FAIL Get Location for Destination Code: " & orig & ". Cannot Print Linehaul DR!", , "Manifest": GoTo EPGP
 End If

 'consignee/shipper
 oy = 5.2: y = 0
 lprnt 0.03125, 0.8125 + oy + y, rsb(0)        'company
 lprnt 0.03125, 1# + oy + y, rsb(1)            'addr
 lprnt 0.03125, 1.1875 + oy + y, rsb(2)        'city
 lprnt 2.25 + 0.03125, 1.1875 + oy + y, rsb(3) 'prov
 lprnt 2.75 + 0.03125, 1.1875 + oy + y, rsb(4) 'pcode/zip
 lprnt 2.75 + 0.03125, 0.8125 + oy + y, rsb(5) 'tel
 y = 0.8125
 lprnt 0.03125, 0.8125 + oy + y, rs(0)
 lprnt 0.03125, 1# + oy + y, rs(1)
 lprnt 0.03125, 1.1875 + oy + y, rs(2)
 lprnt 2.25 + 0.03125, 1.1875 + oy + y, rs(3)
 lprnt 2.75 + 0.03125, 1.1875 + oy + y, rs(4)
 lprnt 2.75 + 0.03125, 0.8125 + oy + y, rs(5)
 
 rs.Close: rsb.Close
 
 Printer.FontItalic = False
 
 lprnt 5.59375, 0.6125 + oy, Format$(mdte, "YYYY-MM-DD") 'date
 lprnt 3.90625, 0.6125 + oy, unp  'unit
 lprnt 3.90625, 0.8125 + oy, "Trailer: " & trp
 
 lprnt 1.9375, 0.5 + oy, l_or  'orig
 lprnt 2.28125, 0.5 + oy, l_de 'DEST
 
 lprnt 0.84375, 2.78125 + oy, "Trip: " & trip
 lprnt 0.84375, 2.9375 + oy + 0.15625, "Seal: " & seal
 lprnt 0.84375, 2.9375 + oy + (3 * 0.15625), "Bills: " & tsh
 lprnt 0.84375, 2.9375 + oy + (5 * 0.15625), " LTL  /  FL"
 If Val(twt) > 0 Then cprnt 5.40625 + (6# - 5.40625) / 2#, 2.9375 + oy + (3 * 0.15625), twt 'weight
 
EPGP:
 Printer.EndDoc
 
 If gateprn <> "" Or inbondprn <> "" Then
   SetDefaultPrinter defprn 'set default printer back to original
   Call SendNotifyMessage(HWND_BROADCAST, WM_WININICHANGE, 0, ByVal "windows")
   Set Printer = odefprn
 End If
 
 Printer.Orientation = 2 'reset default printer to landscape

End Sub


Private Sub PrntMnfstBkGrnd(typ%, cpy%) 'manifest/linehaul, multi-copy#
 Dim j As Integer
 
 Printer.PaintPicture pic1, 1.25, 0.05, 2.05, 0.6154  'speedytransport w x h aspect: 3.25
 
 Printer.Font = Arial: Printer.FontSize = 11: Printer.FontBold = True
 Select Case typ
   Case 0 'delivery manifest
     lprnt 10.4 - Printer.TextWidth("DELIVERY MANIFEST"), 0.145, "DELIVERY MANIFEST"
     lprnt 10.4 - Printer.TextWidth("MANIFESTE DE LIVRAISON"), 0.33, "MANIFESTE DE LIVRAISON"
   Case 1 'linehaul
     lprnt 10.4 - Printer.TextWidth("LINEHAUL MANIFEST"), 0.145, "LINEHAUL MANIFEST"
     lprnt 10.4 - Printer.TextWidth("MANIFESTE DE LONGUE DISTANCE"), 0.33, "MANIFESTE DE LONGUE DISTANCE"
   Case 2 'storage trailer
     lprnt 10.4 - Printer.TextWidth("STORAGE TRAILER MANIFEST"), 0.145, "STORAGE TRAILER MANIFEST"
     lprnt 10.4 - Printer.TextWidth("MANIFESTE DE REMORQUE DE STOCKAGE"), 0.33, "MANIFESTE DE REMORQUE DE STOCKAGE"
 End Select
 
 '{@}
 For j = 0 To 11 'grey boxes for attributes, special instructions, etc.
   Printer.Line (0.9375, 1.4375 + 0.305 + (j * 0.4625))-Step(10# - 0.9375, 0.1515), &HE0E0E0, BF
   If j < 11 Then Printer.Line (0, 1.4375 + 0.4625 + (j * 0.4625))-Step(10.4, 0) 'line below pro#
 Next j
 
 'For j = 0 To 11 'vertical lines
 '  Printer.Line (2.25, 1.4375 + (j * 0.4625))-Step(0, 0.305), HF0F0F0 'shipper left line
 'Next j
 Printer.Line (2.25, 1.4375)-Step(0, 5.3125 + 0.25), HF0F0F0  'shipper left line
 
 For j = 0 To 11
   Printer.Line (6.4375, 1.4375 + (j * 0.4625))-Step(0, 0.3125), HF0F0F0 'cons address city prov
 Next j
 For j = 0 To 11
   Printer.Line (7.125, 1.4375 + (j * 0.4625))-Step(0, 0.3125), HF0F0F0 'COD
 Next j
 For j = 0 To 11
   Printer.Line (7.5625, 1.4375 + (j * 0.4625))-Step(0, 0.3125), HF0F0F0 'PCS
 Next j
 For j = 0 To 11
   Printer.Line (8.1875, 1.4375 + (j * 0.4625))-Step(0, 0.3125), HF0F0F0 'Wt
 Next j
 For j = 0 To 11
   Printer.Line (8.8125, 1.4375 + (j * 0.4625))-Step(0, 0.3125), HF0F0F0 'Cube
 Next j
 For j = 0 To 11
   Printer.Line (9.5, 1.4375 + (j * 0.4625))-Step(0, 0.3125), HF0F0F0 'Appt time
 Next j
 For j = 0 To 11
   Printer.Line (10#, 1.4375 + (j * 0.4625))-Step(0, 0.4625), HF0F0F0 'Dlvy Attrib
 Next j
  
 'section boxes
 bx = 0.4
 Printer.Line (0, 0.75)-(9.5 + bx + 0.5, 7#), , B 'main box
 Printer.Line (0, 7#)-(9.5 + bx + 0.5, 7.375 + 0.25), , B 'bottom box
 Printer.Line (0, 0.75 + 0.375)-(9.5 + bx + 0.5, 0.75 + 0.375 + 0.3125), , BF 'main box column titles black box
 
 'copy-to boxes
 Printer.Line (0.6, 0.17)-(0.71, 0.29), , B 'original
 Printer.Line (0.6, 0.37)-(0.71, 0.49), , B 'office
 Printer.Line (0.6, 0.57)-(0.71, 0.69), , B 'driver
 
 Printer.DrawWidth = 24
 Select Case cpy 'if 3-copy set then 'x'-in appropriate box
   Case 1
     Printer.Line (0.6, 0.17)-(0.71, 0.29): Printer.Line (0.71, 0.17)-(0.6, 0.29)
   Case 2
     Printer.Line (0.6, 0.37)-(0.71, 0.49): Printer.Line (0.71, 0.37)-(0.6, 0.49)
   Case 3
     Printer.Line (0.6, 0.57)-(0.71, 0.69):: Printer.Line (0.71, 0.57)-(0.6, 0.69)
 End Select
 Printer.DrawWidth = 3
 
 'top box separator lines
 Printer.Line (1.4375, 0.75)-Step(0, 0.375) 'left date line
 Printer.Line (3!, 0.75)-Step(0, 0.375)     'driver line
 Printer.Line (4.125, 0.75)-Step(0, 0.375)  'unit line
 Printer.Line (5.5625, 0.75)-Step(0, 0.375) 'trailer line
 Printer.Line (7.25, 0.75)-Step(0, 0.375)   'loaded by
 Printer.Line (9.5, 0.75)-Step(0, 0.375)   'checked by
 
 'main section column lines
 Printer.Line (0.9375, 1.4375)-Step(0, 5.3125 + 0.25), HF0F0F0 'pronumber left line
 'main section column lines white
 Printer.Line (0.9375, 1.4375)-Step(0, -0.6875), &HFFFFFF 'pronumber left line
 Printer.Line (2.25, 1.4375)-Step(0, -0.6875), &HFFFFFF 'shipper
 Printer.Line (3.5625, 1.4375)-Step(0, -0.6875), &HFFFFFF 'cons
 Printer.Line (4.875, 1.4375)-Step(0, -0.6875), &HFFFFFF 'address
 Printer.Line (6.125, 1.4375)-Step(0, -0.6875), &HFFFFFF 'city
 Printer.Line (6.4375, 1.4375)-Step(0, -0.6875), &HFFFFFF 'prov
 Printer.Line (7.125, 1.4375)-Step(0, -0.6875), &HFFFFFF  'COD/Collect
 Printer.Line (7.5625, 1.4375)-Step(0, -0.6875), &HFFFFFF  'PCS
 Printer.Line (8.1875, 1.4375)-Step(0, -0.6875), &HFFFFFF  'Wt
 Printer.Line (8.8125, 1.4375)-Step(0, -0.6875), &HFFFFFF  'Rate
 Printer.Line (9.5, 1.4375)-Step(0, -0.6875), &HFFFFFF 'Appt Time
 Printer.Line (10#, 1.4375)-Step(0, -0.6875), &HFFFFFF 'Dlvy Attrib
 
 'bottom section lines
 Printer.Line (2.5, 7#)-Step(0, 0.625)  'driver sig
 Printer.Line (4.3125, 7#)-Step(0, 0.625)  'tot COD
 Printer.Line (6.0625, 7#)-Step(0, 0.625)  'tot cash
 Printer.Line (7.8125, 7#)-Step(0, 0.625)  'tot wt
 
 'time entry lines
 bx = 0.12 '   0.22
 Printer.Line (1.125, 7.775 + bx)-Step(1.01, 0) 'start time line
 Printer.Line (3.3125, 7.775 + bx)-Step(1.01, 0) 'finish time line
 Printer.Line (5.6875, 7.775 + bx)-Step(1.01, 0) 'lunch time line
 Printer.Line (8.25, 7.775 + bx)-Step(1.01, 0) 'total time line
 bx = 0  '     0.1
 
 'test KIDY directive
 Printer.FontSize = 10
 If dest = "34" Then cprnt 5.15, 7.9, "*  * ** Drop in Doors 2-10 or 32-41 ** *  *"
 
 'time entry text
 Printer.FontSize = 10
 lprnt 1.125 - 0.125 - Printer.TextWidth("START TIME"), 7.63 + bx, "START TIME"
 lprnt 1.125 - 0.125 - Printer.TextWidth("DEBUT"), 7.75 + bx, "DEBUT:"
 lprnt 3.3125 - 0.125 - Printer.TextWidth("FINISH TIME"), 7.63 + bx, "FINISH TIME"
 lprnt 3.3125 - 0.125 - Printer.TextWidth("FIN"), 7.75 + bx, "FIN:"
 lprnt 5.6875 - 0.125 - Printer.TextWidth("LUNCH/BREAKS"), 7.63 + bx, "LUNCH/BREAKS"
 lprnt 5.6875 - 0.125 - Printer.TextWidth("DINER/PAUSE"), 7.75 + bx, "DINER/PAUSE:"
 lprnt 8.25 - 0.125 - Printer.TextWidth("TOTAL HOURS"), 7.63 + bx, "TOTAL HOURS"
 lprnt 8.25 - 0.125 - Printer.TextWidth("HEURES TOTALE"), 7.75 + bx, "HEURES TOTAUX:"

 
 'top-left office copy title-boxes
 lprnt 0.55 - Printer.TextWidth("Original"), 0.15, "Original"
 lprnt 0.55 - Printer.TextWidth("Office"), 0.35, "Office"
 lprnt 0.55 - Printer.TextWidth("Driver"), 0.55, "Driver"
 
 'all other section/column title text starting with all black 7-font text
 Printer.FontSize = 7
 cprnt 1.4375 / 2!, 0.75 + 0.03, "DATE"
 cprnt 1.4375 + ((3! - 1.4375) / 2!), 0.75 + 0.03, "DRIVER / CONDUCTEUR"
 cprnt 3! + ((4.125 - 3!) / 2!), 0.75 + 0.03, "UNIT / UNITE"
 ''cprnt 4.125 + ((5.5625 - 4.125) / 2!), 0.75 + 0.03, "TRAILER / REMORQUE"
 cprnt 2.5 + ((4.3125 - 2.5) / 2!), 6.75 + 0.03 + 0.25, "TOTAL C.O.D. / CR TOTAL"
 cprnt 4.3125 + ((6.0625 - 4.3125) / 2!), 6.75 + 0.03 + 0.25, "TOTAL CASH / COMPTANT TOTAL"
 cprnt 6.0625 + ((7.8125 - 6.0625) / 2!), 6.75 + 0.03 + 0.25, "TOTAL WEIGHT / POIDS TOTAL"
 cprnt 7.8125 + ((9.5 - 7.8125) / 2!), 6.75 + 0.03 + 0.25, "TOTAL CHARGE / FRAIS TOTAUX"
 
 Printer.FontSize = 6.5
 cprnt 2.5 / 2!, 6.75 + 0.03 + 0.25, "DRIVER'S SIGNATURE / SIGNATURE DU CONDUCTEUR"
 
 Printer.FontSize = 5
 cprnt 4.125 + ((5.5625 - 4.125) / 2!), 0.75 + 0.01, "TRAILER-SEAL / REMORQUE-SCEAU"
 cprnt 5.5625 + ((7.25 - 5.5625) / 2!), 0.75 + 0.01, "LOADED BY"
 cprnt 5.5625 + ((7.25 - 5.5625) / 2!), 0.75 + 0.08, "RESPONSABLE DU CHARGEMENT"
 cprnt 7.25 + ((9.5 - 7.25) / 2!), 0.75 + 0.01, "CHECKED BY"
 cprnt 7.25 + ((9.5 - 7.25) / 2!), 0.75 + 0.08, "RESPONSABLE DE LA VERIFICATION"

 'print white column titles text
 Printer.FontSize = 7: Printer.ForeColor = &HFFFFFF
 cprnt 0.9375 / 2!, 1.125 + 0.05, "PRO NO."
 cprnt 0.9375 / 2!, 1.125 + 0.15, "N   DE FACTURE"
 cprnt 0.9375 + ((2.25 - 0.9375) / 2!), 1.125 + 0.05, "SHIPPER"
 cprnt 0.9375 + ((2.25 - 0.9375) / 2!), 1.125 + 0.15, "EXPEDITEUR"
 cprnt 2.25 + ((3.5625 - 2.25) / 2!), 1.125 + 0.05, "CONSIGNEE"
 cprnt 2.25 + ((3.5625 - 2.25) / 2!), 1.125 + 0.15, "CONSIGNATAIRE"
 cprnt 3.5625 + ((4.875 - 3.5625) / 2!), 1.125 + 0.05, "ADDRESS"
 cprnt 3.5625 + ((4.875 - 3.5625) / 2!), 1.125 + 0.15, "ADRESSE"
 cprnt 4.875 + ((6.125 - 4.875) / 2!), 1.125 + 0.05, "CITY"
 cprnt 4.875 + ((6.125 - 4.875) / 2!), 1.125 + 0.15, "VILLE"
 Printer.FontSize = 6.5
 cprnt 6.125 + ((6.4375 - 6.125) / 2!), 1.125 + 0.1, "PROV"
 Printer.FontSize = 7
 cprnt 7.125 + ((7.5625 - 7.125) / 2!), 1.125 + 0.05, "PCS"
 cprnt 7.125 + ((7.5625 - 7.125) / 2!), 1.125 + 0.15, "QTE"
 cprnt 7.5625 + ((8.1875 - 7.5625) / 2!), 1.125 + 0.05, "WEIGHT"
 cprnt 7.5625 + ((8.1875 - 7.5625) / 2!), 1.125 + 0.15, "POIDS"
 cprnt 8.1875 + ((8.8125 - 8.1875) / 2!), 1.125 + 0.05, "CUBE WT"
 cprnt 8.1875 + ((8.8125 - 8.1875) / 2!), 1.125 + 0.15, "CUBE"
 cprnt 9.5 + ((10# - 9.5) / 2!), 1.125 + 0.05, "RATE"
 cprnt 9.5 + ((10# - 9.5) / 2!), 1.125 + 0.15, "TAUX"
 cprnt 10# + ((10.4 - 10#) / 2!), 1.125 + 0.05, "HRS"
 cprnt 10# + ((10.4 - 10#) / 2!), 1.125 + 0.15, "RCVG"
 
 Printer.FontSize = 5
 cprnt 6.4375 + ((7.125 - 6.4375) / 2!), 1.125 + 0.01, "COD OR CASH"
 cprnt 6.4375 + ((7.125 - 6.4375) / 2!), 1.125 + 0.08, "COLLECT"
 cprnt 6.4375 + ((7.125 - 6.4375) / 2!), 1.125 + 0.15, "CR OU COMPTANT."
 cprnt 6.4375 + ((7.125 - 6.4375) / 2!), 1.125 + 0.22, "PORT DU"
 cprnt 8.8125 + ((9.5 - 8.8125) / 2!), 1.125 + 0.01, "APPT."
 cprnt 8.8125 + ((9.5 - 8.8125) / 2!), 1.125 + 0.08, "TIME"
 cprnt 8.8125 + ((9.5 - 8.8125) / 2!), 1.125 + 0.15, "TEMPS"
 cprnt 8.8125 + ((9.5 - 8.8125) / 2!), 1.125 + 0.22, "CONVENU"
 'write 'o' in 'No DE FACTEUR'
 Printer.FontSize = 5
 lprnt 0.148, 1.125 + 0.143, "o"
 
 Printer.ForeColor = &H0
 
 Printer.Orientation = 2
 'print barcode center-top
 Printer.PaintPicture barc(0).Image, 5.25 - 1.01, 0.25, 2#, 0.25  'x1, y1, Width, Height
 Printer.Orientation = 2

End Sub

Private Sub PrntMBOL(objw!)
 Dim pcnt As Integer
 Dim prnerr As Boolean
 'printer already scaled to inches
 On Error Resume Next
 obj.PrintWidth = objw: obj.PrintLeft = 0: obj.PrintTop = 0 'no margins, use entire printable area
MRETPRNT:
 result = obj.PrintFile 'send job to printer
 If result <> 0 Then 'print error if scaled outside of printer workable margins
   pcnt = pcnt + 1
   If pcnt < 5 Then  'reduce printwidth by up to 1 inch in 1/4" increments
     objw = objw - 0.25: obj.PrintWidth = objw: GoTo MRETPRNT
   End If
   prnerr = True 'print error due to some other reason
 End If
 If prnerr Then MsgBox "Error: " & result & "  Could not print to Default Printer.", , "Manifest"
End Sub

Private Sub c_tr_Click()
 Dim j As Integer
 Dim dm As String
 Select Case Trim$(t_tro)
   Case "", "SP", "SDT" 'Speedy sdt*
     If trailer <> "" And trailer = unit Then 'straight or cube - must change unit b4 trailer can change
       frun.Visible = True: lst_un.SetFocus: fr1.Enabled = False
     Else
       frtr.Visible = True: lst_tr.SetFocus: fr1.Enabled = False
     End If
   Case "HOL"
     If Not uholtr Then 'not already displaying Holland
       uholtr = True: uexltr = False
       lst_htr.Clear
       For j = 0 To UBound(holtr)
         If holtr(j, 0) = "" Then Exit For
         lst_htr.AddItem holtr(j, 0) & "  " & holtr(j, 1)
       Next j
     End If
     frhtr.Visible = True: lst_htr.SetFocus: fr1.Enabled = False
   Case "EXL"
     If Not uexltr Then 'not already displaying Estes
       uexltr = True: uholtr = False
       lst_htr.Clear
       For j = 0 To UBound(exltr)
         If exltr(j, 0) = "" Then Exit For
         lst_htr.AddItem exltr(j, 0) & "  " & exltr(j, 1) & "  " & exltr(j, 2)
       Next j
     End If
     frhtr.Visible = True: lst_htr.SetFocus: fr1.Enabled = False
   Case Else
     MsgBox "Trailer List for Owner Code " & t_tro & " is Not Available.", , "Manifest"
 End Select
End Sub

Private Sub c_un_Click()
 Select Case Trim$(t_uno)
   Case "", "SP", "SDT" 'sdt*
     frun.Visible = True
     If lst_un.Visible Then lst_un.SetFocus
     fr1.Enabled = False
   Case Else
     MsgBox "Unit List for Owner Code " & t_uno & " is Not Available.", , "Manifest"
 End Select
End Sub

Private Sub dpro_Click() 'popup menu - view probill
 pch = True
 vpro = g3.TextMatrix(g3y, 1)
 prob.Show 1
 pro.Visible = True: pro.SetFocus
 pch = False
End Sub

Private Function GetDefPrn() As String 'get current Window default printer name
 Dim dm As String
 Dim n As Long
 GetDefaultPrinter vbNullChar, n 'get length of current default printer name string
 If n = 0 Then
   MsgBox "FAIL Get Default Printer. Aborting . . ."
   GetDefPrn = "": Exit Function
 End If
 dm = Space$(n) 'build empty string of this length
 GetDefaultPrinter dm, n 'fill empty string with printer name
 GetDefPrn = Left$(dm, InStr(dm, vbNullChar) - 1) 'extract current default printer name
End Function

Private Sub Form_Unload(Cancel As Integer)
 Set ogateprn = Nothing: Set oinbondprn = Nothing
 Set fo = Nothing
 If rs.State <> 0 Then rs.Close
 Set rs = Nothing
 If rsb.State <> 0 Then rsb.Close
 Set rsb = Nothing
 If rsc.State <> 0 Then rsc.Close
 Set rsc = Nothing: Set cmd = Nothing
 conn.Close: Set conn = Nothing
 If ir.State <> 0 Then ir.Close
 Set ir = Nothing
 iconn.Close: Set iconn = Nothing
End Sub

Private Sub Form_Load()

 Dim dark As Boolean, thin As Boolean, destonly As Boolean
 Dim i As Integer, j As Integer, n As Integer
 Dim dum As String * 255
 Dim dm As String, dm2 As String
 Dim s() As String 'for Split Fn
 
 remoteproedb = False
 dm = Command()
 If dm <> "" Then
   n = InStr(dm, "|")
   If n > 0 Then
     uinits = UCase$(Left$(dm, n - 1))
     j = InStr(dm, "[")
     If j = 0 Then
       upwrd = Mid$(dm, n + 1)
     Else
       If j > n + 1 Then upwrd = Mid$(dm, n + 1, j - (n + 1))
       remoteproed = Mid$(dm, j + 1)
       Select Case Left$(remoteproed, 4)
         Case "EXLA", "WARD"
           If Len(remoteproed) = 14 Then remoteproedb = True Else End
         Case "HMES"
           Select Case Len(remoteproed)
             Case 13, 14: remoteproedb = True
             Case Else: End
           End Select
         Case "PYLE"
           If Len(remoteproed) = 13 Then remoteproedb = True Else End
         Case "SEFL"
           Select Case Len(remoteproed)
             Case 11, 12: remoteproedb = True
             Case Else: End
           End Select
         Case Else
           If Val(remoteproed) > 9999 Then remoteproedb = True Else End
       End Select
     End If
   Else
     uinits = dm
   End If
   ucmdl = True 'indicate user inits/password passed thru commandline argument at program start
 End If
 
 sdir = "C:\ProTrace\StagingFolders\manifest\"
 Set fo = CreateObject("Scripting.FileSystemObject") 'init file i/o ops
 
 trailerloc.Path = "C:\ProTrace\TrailerLastLocationImage"
 trailerloc.Pattern = "*.jpg"
 trailerloc.Refresh

 'get Speedy Ops client db connection string from 'ini' file
 n = GetPrivateProfileString("connstr", ByVal "0", "", dum, 255, sdir & "TRPMNFST.INI")
 If n = 0 Then
   MsgBox "Unable to Load Ops Connection String! Program Abort"
   End
 End If
 'establish connection with ADO db objects (global scope)
 conn.Open (Left$(dum, n)) '(ConnectString)
 rs.ActiveConnection = conn  'primary SQL query recordset/command object
 
 If remoteproedb Then
   rs.Open "select pronumber from probill where pronumber='" & remoteproed & "'"
   If rs.EOF Then
     rs.Close: End
   End If
   rs.Close
 End If
 
 rsb.ActiveConnection = conn 'lookup/x-ref query recordset
 rsc.ActiveConnection = conn '  "      "     "       "
 pb.ActiveConnection = conn
 cmd.ActiveConnection = conn 'SQL command object
 cmd.CommandType = adCmdText 'slight speed increase - object does not have to pre-examine query
 
 gpissuer = "N/A"
 If ucmdl Then
   rs.Open "select terminal,rights,manifestpostedit,nam,emptygatepass from uinits where init='" & uinits & "'"
   If Not rs.EOF Then
     uterm = rs(0): urghts = rs(1): gpissuer = Trim$(rs(3))
     If Len(gpissuer) < 2 Then gpissuer = "N/A"
     If rs(2) = 1 Then mef = True '//^\\_//^\\
     ''If rs(4) = 1 Then c_pegp.Visible = True  'no access rights needed for now
   End If
   rs.Close
 End If
 If mef Then '//^\\_//^\\
   o_prnt(2).Visible = True: l_prnt(2).Visible = True
 End If
  
 'get Speedy Images client db connection string from 'ini' file
 n = GetPrivateProfileString("iconnstr", ByVal "0", "", dum, 255, sdir & "TRPMNFST.INI")
 If n = 0 Then
   MsgBox "Unable to Load Images Connection String! Program Abort"
   End
 End If
 'establish connection with ADO db objects (global scope)
 iconn.Open (Left$(dum, n)) '(ConnectString)
 ir.ActiveConnection = iconn  'primary SQL query recordset/command object
 icmd.ActiveConnection = iconn: icmd.CommandType = adCmdText
    
 'get Speedy Correspondence db connection string from 'ini' file
 n = GetPrivateProfileString("aconnstr", ByVal "0", "", dum, 255, sdir & "TRPMNFST.INI")
 If n = 0 Then
   MsgBox "Unable to Load Correspondence Connection String! Program Abort"
   End
 End If
 aconnstr = Left$(dum, n)

 On Error Resume Next
 
 Err = 0

 'units list - ProTrace via 'Collective Data' API transfers to mySQL via Evoship service
 dm = "select num,type,assignment from trucks where owner='SZTG' and active='1' " & _
      "and num <> 'XTRUCK' " & _
      "order by type desc, cast(num as binary) desc"
 rs.Open dm
 If rs.EOF Then
   rs.Close: MsgBox "FAIL Retrieve Unit List! Aborting . .", , "Manifest": End
 End If
 lst_un.Clear: lst_un.AddItem "    - UN-ASSIGNED -"
 lst_un.AddItem " 6000" & "  xTOR      CPRail"
 lst_un.AddItem " 6001" & "  xMTL      CPRail"
 
 Dim test As String
 Do
   Select Case rs(0)
     Case "XTRUCK"
     Case Else
       s = Split(rs(1), " ")
       dm2 = Left$(Trim$(s(0)), 8)
       If InStr(UCase$(rs(1)), "TAILGATE") > 0 Or InStr(UCase$(rs(1)), " TG") > 0 Then dm2 = dm2 & " TG"
       lst_un.AddItem Space(8 - Len(rs(0))) & rs(0) & "  " & dm2 & Space(11 - Len(dm2)) & "  " & Replace(rs(2), " ", ""): rs.MoveNext
   End Select
 Loop Until rs.EOF
 rs.Close: frun.Top = 555: frun.Left = 1470 '2175
 g2pro.Left = 75
 
 'trailers list
 dm = "select num,length,type from trailers where owner='SZTG' and active='1' " & _
      "and num not in ('XTRAILERS','XHOLLAND','XESTES') " & _
      "order by cast(num as binary)"
 rs.Open dm
 If rs.EOF Then
   rs.Close: MsgBox "FAIL Retrieve Trailer List! Aborting . .", , "Manifest": End
 End If
 'write special, non-speedy, etc. trailers before writing main SZTG trailers list
 lst_tr.AddItem "    - UN-ASSIGNED -"
 'get all service trailers into working array - used to put trailer service codes alongside items in main list
 Erase sv
 dm = "select trailer,owner,svccode from trlrsvc where owner='SZTG' order by owner,trailer"
 rsb.Open dm
 If Not rsb.EOF Then
   j = 0
   Do
     j = j + 1: ReDim Preserve sv(1 To j): sv(j).t = rsb(0): sv(j).o = rsb(1): sv(j).c = rsb(2)
     rsb.MoveNext
   Loop Until rsb.EOF
 End If
 rsb.Close
 'insert special, non-Speedy, etc. trailers at the list start
 rsb.Open "select trailer,descrip,length from trip_sptrailer where active='1' order by trailer"
 If Not rsb.EOF Then
   Do
     dm2 = "  "
     For j = 1 To UBound(sv)
       If rsb(0) = sv(j).t Then
         dm2 = sv(j).c: Exit For
       End If
     Next j
     lst_tr.AddItem dm2 & Space(11 - Len(rsb(0))) & rsb(0) & "  " & Format$(Val(rsb(2)), "00") & "  " & rsb(1)
     rsb.MoveNext
   Loop Until rsb.EOF
 End If
 rsb.Close
'now write Speedy trailers to list
 Do
   dm = rs(0): dm2 = "  "
   For j = 1 To UBound(sv) 'look for this trailer in service codes list
     If rs(0) = sv(j).t Then 'if found will put service code alongside this trailer below in 'additem'
       dm2 = sv(j).c: Exit For
     End If
   Next j
   lst_tr.AddItem dm2 & Space(11 - Len(dm)) & dm & "  " & Format$(Val(rs(1)), "00") & "  " & Replace(rs(2), "-", "")
   rs.MoveNext
 Loop Until rs.EOF
 rs.Close
 
SKPSCC:
 
 frtr.Top = 945: frtr.Left = 1545 '1635 '2175 '1545 '@#@ 2175
 
 'Holland trailers list
 dm = "select trailerno, licence from paps_hol_trailers where cast(trailerno as unsigned) between '100000' and '999999' order by cast(trailerno as unsigned)"
 rs.Open dm
 If rs.EOF Then
   rs.Close: MsgBox "FAIL Retrieve Holland Trailer List!", , "Manifest"
 Else
  j = -1
   Do
     j = j + 1: holtr(j, 0) = rs(0): holtr(j, 1) = rs(1): rs.MoveNext 'buff trailer info to array
   Loop Until rs.EOF
   rs.Close
 End If
 
 'EXLA trailers list
 dm = "select trailer,licence,length from paps_pars_exla_trailers order by length,trailer"
 rs.Open dm
 If rs.EOF Then
   rs.Close: MsgBox "FAIL Retrieve Estes Trailer List!", , "Manifest"
 Else
  j = -1
   Do
     j = j + 1: exltr(j, 0) = rs(0): exltr(j, 1) = rs(1): exltr(j, 2) = rs(2): rs.MoveNext 'buff trailer info to array
   Loop Until rs.EOF
   rs.Close
 End If
 
 'position partner trailers list box
 frhtr.Top = 900: frhtr.Left = 2175
  
 'populate orig/dest lists
 rs.Open "select origdest,code,tripcode,destcity from linehaul_status_codes where origdest <> '' " & _
         "and cast(code as unsigned) < '1000' order by cast(code as unsigned)"
 If rs.EOF Then
   rs.Close: MsgBox "FAIL Retrieve Linehaul Codes! Aborting . .", , "Manifest"
   End
 End If
 Do
   arg = rs(1)
'   If Val(rs(1)) = 83 Then
'     arg = arg
'   End If
   destonly = False
   Select Case Val(rs(1))
     Case 1: dm2 = "TOR"
     Case 3: dm2 = "MIS"
     Case 4: dm2 = "MIL"
     Case 5: dm2 = "LON"
     Case 6: dm2 = "WIN"
     Case 8: dm2 = "PIC"
     Case 10: dm2 = "BRO"
     Case 11: dm2 = "LAC"  'VAU
     Case 12: dm2 = "MTL"
     Case 14: dm2 = "QUE": destonly = True ' "QCY"
     Case Else: dm2 = rs(0)
   End Select
   dm = dm2 & Space(7 - Len(dm2)) & rs(1) & Space(4 - Len(rs(1))) & rs(3)
         'dm = rs(0) & Space(7 - Len(rs(0))) & rs(1) & Space(4 - Len(rs(1))) & rs(3)
   If Not destonly Then
     lst_or.AddItem dm: lst_or.ItemData(lst_or.NewIndex) = rs(1)
   End If
   lst_de.AddItem dm: lst_de.ItemData(lst_de.NewIndex) = rs(1)
   lhcod(rs(1)) = rs(2) 'buf tripcode prefix to array
   destcty(rs(1)) = rs(3) 'buf full cityname for tripcode
   rs.MoveNext
 Loop Until rs.EOF
 rs.Close
 lst_or.Top = fr1.Top + c_or.Top + c_or.Height: lst_or.Left = fr1.Left + c_or.Left + c_or.Width - lst_or.Width
 lst_de.Top = fr1.Top + c_de.Top + c_de.Height: lst_de.Left = fr1.Left + c_de.Left + c_de.Width - lst_de.Width

 'outside power owner code lists/buffers
 lst_uno.Top = 345: lst_uno.Left = 765: lst_tro.Top = 735: lst_tro.Left = 765
 'populate lists
 rs.Open "select code, nam from trip_owners where istractor='1' and dispmnfst='1' order by ord"
 If rs.EOF Then
   MsgBox "FAIL Read Outside Unit Owner Codes!", , "Manifest"
 Else
   lst_uno.Clear: n = -1
   Do
     n = n + 1: uownr(n, 0) = rs(0): uownr(n, 1) = rs(1): lst_uno.AddItem rs(0) & Chr(9) & rs(1)
     rs.MoveNext
   Loop Until rs.EOF
 End If
 rs.Close
  '' old 3-char owner codes -> rs.Open "select code,nam,scac from trip_owners where istrailer='1' and tshw='1' order by tord"
 rs.Open "select code,nam,scac from trip_owners where istrailer='2' and tshw='2' order by tord"
 If rs.EOF Then
   MsgBox "FAIL Read Outside Unit Owner Codes!", , "Manifest"
 Else
   lst_tro.Clear: n = -1
   Do
     n = n + 1: tsc(n, 0) = rs(0): tsc(n, 1) = rs(1): tsc(n, 2) = rs(2): lst_tro.AddItem rs(0) & Chr(9) & rs(1)
     rs.MoveNext
   Loop Until rs.EOF
 End If
 rs.Close
 
 'manifest date control
 mdte.MaxDate = DateValue(Now + 14)
 mdte = DateValue(Now)
 mdte_Change
 
 'setup manifest table
 g2.Cols = 24: g2.Rows = 208: j = 0
 g2.WordWrap = True: g2.MergeCells = flexMergeRestrictRows   'Merge TURNS OFF ROW SELECTION!!
 For n = 0 To 207
   'tall row + thin row = shipment (probill, etc. data + account codes)
   g2.Row = n: thin = False: dark = False
   If n > 0 Then
     If n / 2 <> Fix(CSng(n) / 2!) Then thin = True 'odd nos. = thin row
     If n > 1 And (n - 2) / 4 = Fix(CSng(n - 2) / 4!) Then dark = True 'every 4th row after row 2
   End If
   If thin Then
     g2.RowHeight(n) = 195: g2.TextMatrix(n, 18) = "1" 'mark row as 'thin'
     For i = 0 To 2: g2.MergeRow(n) = True: g2.TextMatrix(n, i) = " ": Next i
     For i = 3 To 9 '{@} was for i=1 to 9
       g2.MergeRow(n) = True: g2.TextMatrix(n, i) = " "
     Next i 'control which cells merge in attributes thin-row
   Else
     g2.RowHeight(n) = 585: g2.Col = 0: g2.CellFontName = "Arial"
     Select Case n
       Case Is < 19
         g2.CellFontSize = 8: g2.CellFontBold = True
       Case Is < 198
         g2.CellFontSize = 7: g2.CellFontBold = True
       Case Else
         g2.CellFontSize = 6: g2.CellFontBold = False
     End Select
     j = j + 1: g2.TextMatrix(n, 0) = CStr(j)
   End If
   If dark Then
     For i = 0 To g2.Cols - 1
       g2.Col = i: g2.CellBackColor = &HE0ECED
       g2.Row = n + 1: g2.Col = i: g2.CellBackColor = &HDCE9EA: g2.Row = n
     Next i
     g2.TextMatrix(n, 19) = "1": g2.TextMatrix(n + 1, 19) = "1" 'mark row pair as 'dark'
   End If
   g2.Col = 0: g2.CellAlignment = flexAlignRightCenter: g2.Col = 1: g2.CellAlignment = flexAlignCenterCenter
   g2.Col = 2: g2.CellAlignment = flexAlignLeftCenter: g2.Col = 3: g2.CellAlignment = flexAlignLeftCenter
   g2.Col = 4: g2.CellAlignment = flexAlignRightCenter: g2.Col = 5: g2.CellAlignment = flexAlignCenterCenter
   g2.Col = 6: g2.CellAlignment = flexAlignCenterCenter: g2.Col = 7: g2.CellAlignment = flexAlignCenterCenter
   g2.Col = 8: g2.CellAlignment = flexAlignCenterCenter: g2.Col = 9: g2.CellAlignment = flexAlignLeftCenter
   
   'g2.Col=xx
   If thin Then 'use italics on thin (attributes) row
     ''g2.Col = 1: g2.CellFontName = Arial: g2.CellFontSize = 7: g2.CellFontBold = True
     For i = 0 To 2:
       g2.Col = i: g2.CellAlignment = flexAlignLeftCenter: g2.MergeRow(n) = True: g2.TextMatrix(n, i) = " "
     Next i
     For i = 3 To 11 '{@} was 0 to 11
       g2.Col = i: g2.CellFontItalic = True
     Next i
   End If
 Next n
 g2.ColWidth(0) = 255 '2 char + gridline
 g2.ColWidth(1) = 1065 + 405 + 45 '10 pro/date/PCTN ~~
 g2.ColWidth(2) = 1620 '18 shipper
 g2.ColWidth(3) = 2790 '30 consignee name/address/city-prov-postzip
 g2.ColWidth(4) = 630  '6 pcs/type
 g2.ColWidth(5) = 505  '5 totwt
 g2.ColWidth(6) = 1230 'appt date/time
 g2.ColWidth(7) = 900  'rcvg hrs
 g2.ColWidth(8) = 870  'COD/Collect
 g2.ColWidth(9) = 1665 'last 3 statuses
 g2.ColWidth(10) = 195 'BOL available = '*' char (checks Speedy & STPC BOL's)
 'hidden cols
 '11 = image index (see 15)
 '12 = If = 1 then Early Delivery
 '13 = If = 1 then Appointment
 '14 = If > 0 then Cube Wt
 '15 = BOL source - 1 = std BOL scan from img_sfmaster, 2 = STPC BOL fax/scan from img_stpc
 '16 - Hazmat indicator
 '17 - pronumber
 '18 - 0=thick row, 1=thin row
 '19 - alternate light/dark row-pairs where 1=dark
 '20 - agent/partner pronumber
 '21 - freight dimensions (cubevol)
 '22 - delivery terminal 01Nov13(DNY)
 '23 - cons prov 22Aug18(DNY) re: Tumi routing Wstrn Cda - see DispProRow for write

 ChkPrntr 'look for user's default printer (set global var 'prnavail' True if found) '21Nov13 - add look for 'GatePassPrinter' & set global var
 If defprn = "GatePassPrinter" Then
   dm = "Your Default Windows Printer is set to 'GatePassPrinter'." & vbCrLf & vbCrLf & _
        "Please Do Not Print Manifests on This Printer unless absolutely necessary." & vbCrLf & vbCrLf & _
        "   * Select 'OK' to End and Reset your Default Windows Printer *"
   If MsgBox(dm, vbDefaultButton1 + vbInformation + vbOKCancel, "Default Windows Printer Verification") = vbOK Then End
   gatepassprn = False: gateprn = "": Set ogateprn = Nothing
 ElseIf defprn = "InlandInbondPrinter" Then
   dm = "Your Default Windows Printer is set to 'InlandInbondPrinter'." & vbCrLf & vbCrLf & _
        "   * Select 'OK' to End and Reset your Default Windows Printer *"
   If MsgBox(dm, vbDefaultButton1 + vbInformation + vbOKOnly, "Default Windows Printer Verification") = vbOK Then End
 Else
   If gateprn = "GatePassPrinter" Then lpg.Visible = True
 End If
 Set obj = New PictPlus60Vic.clsPicturePlus 'activate PictPlus object for BOL printing
    
 l_arr(0) = "": l_arr(1) = "": l_od = "": l(46) = Format$(Now, "DDD DD-MMMM-YYYY")
 oldrow = -1: curow = 0: tlst.Clear          'file transfer list
 
 If ucmdl Then
   init = uinits: init.TabStop = False: init.Locked = True: init_Validate False
 End If
 
 InitProDDE
 
 'set base size for multi-purpose picture box (eg. printing barcodes & images)
 barc0w = 6030: barc0h = 810: drsdir = ""
 'set path to ProTrace base logo images
 If fo.FileExists("C:\Program Files\PROTRACE\trcrnrbk.JPG") Then
   drsdir = "C:\Program Files\PROTRACE\"
 Else
   If fo.FileExists("C:\Program Files(x86)\PROTRACE\trcrnrbk.JPG") Then
     drsdir = "C:\Program Files(x86)\PROTRACE\"
   End If
 End If
 
 '14Oct14 init AutoMail emailer vars
 If fo.FileExists("c:\temp\stpc\am\automail.exe") Then mailsendon = True Else mailsendon = False
 aempth = "c:\temp\stpc\am\": aminipth = aempth & "automail.ini": aempoll = "c:\temp\stpc\automail\"
 aemfrm = "": aemcc = "": aembcc = ""

 intv = 15
 tmr.Enabled = True  'start data transfer tracking 2-sec timer
 tmr_Timer
 
 If ucmdl And remoteproedb Then
   t_fpro = remoteproed
   c_fgo_Click
   For j = 0 To g2.Rows - 1 Step 2
     If g2.TextMatrix(j, 17) = remoteproed Then
       If j > 0 Then HiRo 0, 0, 0
       HiRo j, 1, 0
       Exit For
     End If
   Next j
   c_prnt.Enabled = False: c_new.Enabled = False: c_ldtrp.Enabled = False: c_fgo.Enabled = False
   c_lddte.Enabled = False: c_clr.Enabled = False: c_prntbols.Enabled = False: c_prnttom.Enabled = False
   c_zstat.Enabled = False: mdte.Enabled = False: g2pro.Enabled = False: g2pop.Enabled = False
   t_fpro = ""
 End If
 ucmdl = False
 If Len(init) = 3 Then
   t_tr.TabIndex = 0: t_tr.SetFocus
 End If

End Sub

Private Sub Form_KeyDown(KeyCode%, Shf%)
 If KeyCode = vbKeyF10 Then KeyCode = vbKeyF11 're-direct Windows form key to user key
End Sub

Private Sub c_initdde_Click()
 InitProDDE
End Sub

Private Sub InitProDDE()
 On Error Resume Next
 corrpro.LinkTopic = "protrace_cs|attchfrm": corrpro.LinkItem = "corrprocpy": corrpro.LinkMode = vbLinkNotify
 If Err = 0 Then
   corrpro.Visible = True
 Else
   corrpro.LinkMode = vbLinkNone: Err = 0: corrpro.LinkTopic = "protrace_cs2|attchfrm": corrpro.LinkItem = "corrprocpy": corrpro.LinkMode = vbLinkNotify
   If Err = 0 Then
     corrpro.Visible = True
   Else
     corrpro.Visible = False: corrpro.LinkMode = vbLinkNone: Err = 0
   End If
 End If
 tttrp.LinkTopic = "trailertrace|main": tttrp.LinkItem = "mlnk1": tttrp.LinkMode = vbLinkNotify
 If Err = 0 Then
   tttrp.Visible = True
 Else
   tttrp.Visible = False: tttrp.LinkMode = vbLinkNone: Err = 0
 End If

End Sub

Private Sub corrpro_LinkNotify()  'DDE link to Pro# changes in Correspondence screen
 corrpro.LinkRequest              ' see InitProDDE and 'corrpro' field
 If t_fpro <> corrpro Then t_fpro = corrpro
 corrpro = ""
End Sub

Private Sub tttrp_LinkNotify()  'DDE link to trip# from TrailerTrace
 tttrp.LinkRequest
 Select Case Val(tttrp)
   Case 0, Is < 100000, Is > 999999: Exit Sub
   Case Else
     If Right(trip, 6) = tttrp Then Exit Sub
 End Select
 If l_ne = "EDIT" Then c_clr_Click
 chksvcno = True: trip = tttrp: c_ldtrp_Click
 chksvcno = False: trip.SetFocus
End Sub

'---- Pop-Up Menu Commands -------------------------------------
Private Sub g2pop1_Click() 'insert blank row
 Dim j As Integer
 Dim dm As String
 If curow = 206 Then Exit Sub
 If g2.TextMatrix(206, 17) <> "" Then Exit Sub
 For j = 204 To curow Step -2
   If g2.TextMatrix(j, 17) = "" Then
     Clrg2 j + 2
   Else
     For k = 1 To 17
       g2.TextMatrix(j + 2, k) = g2.TextMatrix(j, k)
     Next k
     g2.TextMatrix(j + 2, 20) = g2.TextMatrix(j, 20): g2.TextMatrix(j + 2, 21) = g2.TextMatrix(j, 21)
     g2.TextMatrix(j + 2, 22) = g2.TextMatrix(j, 22):  g2.TextMatrix(j + 2, 23) = g2.TextMatrix(j, 23)
     For k = 1 To 8
       g2.TextMatrix(j + 3, k) = g2.TextMatrix(j + 1, k)
     Next k
     g2.TextMatrix(j + 3, 20) = g2.TextMatrix(j + 1, 20): g2.TextMatrix(j + 3, 21) = g2.TextMatrix(j + 1, 21)
     g2.TextMatrix(j + 3, 22) = g2.TextMatrix(j + 1, 22): g2.TextMatrix(j + 3, 23) = g2.TextMatrix(j + 1, 23)
   End If
 Next j
 Clrg2 curow
 g2pro = ""
End Sub

Private Sub g2pop2_Click() 'remove row
 Dim j As Integer, k As Integer
 Dim dm As String
 If g2.TextMatrix(curow, 17) <> "" Then 'warn if row contains shipment
   dm = "CONFIRM Remove Row: " & CStr(curow / 2 + 1) & " Pro: " & g2.TextMatrix(curow, 17)
   If MsgBox(dm, vbDefaultButton2 + vbInformation + vbOKCancel, "Remove Row from Manifest") = vbCancel Then Exit Sub
 End If
 If g2.TextMatrix(curow, 17) <> "" Then Clrg2 curow
 For j = curow + 2 To 206 Step 2
   If g2.TextMatrix(j, 17) <> "" Then
     For k = 1 To 17
       g2.TextMatrix(j - 2, k) = g2.TextMatrix(j, k)
     Next k
     g2.TextMatrix(j - 2, 20) = g2.TextMatrix(j, 20): g2.TextMatrix(j - 2, 21) = g2.TextMatrix(j, 21)
     g2.TextMatrix(j - 2, 22) = g2.TextMatrix(j, 22): g2.TextMatrix(j - 2, 23) = g2.TextMatrix(j, 23)
     For k = 1 To 8
       g2.TextMatrix(j - 1, k) = g2.TextMatrix(j + 1, k)
     Next k
     g2.TextMatrix(j - 1, 20) = g2.TextMatrix(j + 1, 20): g2.TextMatrix(j - 1, 21) = g2.TextMatrix(j + 1, 21)
     g2.TextMatrix(j - 1, 22) = g2.TextMatrix(j + 1, 22): g2.TextMatrix(j - 1, 23) = g2.TextMatrix(j + 1, 23)
     Clrg2 j
   End If
 Next j
 g2pro = g2.TextMatrix(curow, 17)
 WrTotWtCube 'recompute total manifested wt
End Sub

'---- manifest pro list ----------------------------------------

Private Sub g2_MouseDown(Button%, Shf%, x!, y!)
 Dim n As Integer, j As Integer
 Dim b As Boolean
 Dim rw As Variant
 
 'get row clicked
 rw = g2.TopRow / 2 + Fix(y / 780!) + 1 'get row label
 n = (rw - 1) * 2 'set row to 'thick' row of pair clicked
 If n <> curow Then 'select row clicked if different from current
   newrow = n
   If oldrow > -1 Then HiRo oldrow, 0, 0
   HiRo newrow, 1, 0
   oldrow = newrow: g2t = g2.TopRow
 End If
 
 If remoteproedb Then Exit Sub
 
 If Button = 1 And g2pro.Visible Then g2pro.SetFocus 'always move focus to pro entry field
 
 If Button <> 2 Then Exit Sub 'right-click for pop-up menu
 
 'adjust popup menu according to selected row data
 If g2.TextMatrix(n, 17) = "" Then
   g2pop0.Caption = "Row " & CStr(rw) & " - Blank"
   g2pop1.Enabled = True 'insert row
   'check remove row logic
   g2pop2.Enabled = False: g2pop3.Enabled = False: g2pop4.Enabled = False
   g2pop5.Enabled = False: g2pop6.Enabled = False: g2pop7.Enabled = False
   g2pop8.Enabled = False: g2pop9.Enabled = False
   If n < 206 Then
     For j = n + 2 To 206 Step 2 'look for a pro after this line
       If g2.TextMatrix(j, 17) <> "" Then
         g2pop2.Enabled = True: Exit For 'pro found, allow remove blank row
       End If
     Next j
   End If
 Else
   g2pop0.Caption = "Row " & CStr(rw) & " - Pro: " & g2.TextMatrix(n, 17)
   g2pop1.Enabled = True 'insert row
   g2pop2.Enabled = True 'remove row
   g2pop5.Enabled = True 'adjust totwt

   If g2.TextMatrix(n, 11) = "" Then 'if BOL scanimage found
     g2pop3.Enabled = False: g2pop4.Enabled = False
   Else
     g2pop3.Enabled = True: g2pop4.Enabled = True: g2pop5.Enabled = True
   End If
   g2pop7.Enabled = True 'update last location-row
   g2pop8.Enabled = True 'view correspondence
   g2pop9.Enabled = True 'Print DR
 End If
 PopupMenu g2pop
 
End Sub

Private Sub g2pro_GotFocus()
 g2profoc = True
 g2pro.BackColor = &HFFFFFF
 HookMnfst mnfst.hWnd
 g2foc = True
End Sub
Private Sub g2pro_LostFocus()
 g2pro.BackColor = &HD6E6E6
 g2foc = False
 UnHookMnfst
End Sub
Private Sub g2_GotFocus()
 g2profoc = False
 HookMnfst mnfst.hWnd
 g2foc = True
 If Not g2pro.Visible Then g2pro.Visible = True
End Sub
Private Sub g2_LostFocus()
 g2foc = False
 UnHookMnfst
End Sub
'triggered by mousewheel movement (see Module 'MWheel.bas')
Public Sub MouseWheelMnfst(ByVal zDelta As Long)
''NOTE: this routine must NOT be impeded or stop during execution or program will crash
''      because windows message buffer will overflow in ~ 5 secs. on typical P4 machine
 Dim i As Integer
 If g2foc Then 'results grid has focus, use mousewheel
   If g2.TextMatrix(g2.Row, 18) = "1" Then i = 1 Else i = 0 'check if current row is 'thin'
   g2.Row = g2.Row - i 'correct row (thick) to account for pairs
   Select Case zDelta 'check wheel click increment amount (1 click = 120)
     Case Is = 120   'for each mousewheel movement forward (away from user) by 1 click - scroll up
       If curow < 2 Then Exit Sub  'If g2.Row = 0 Then Exit Sub
       oldrow = g2.Row: newrow = oldrow - 2: curow = newrow 'move 2 rows up
       If oldrow > -1 Then HiRo oldrow, 0, 9
       HiRo newrow, 1, 9
       oldrow = newrow
       If newrow < g2.TopRow Then g2.TopRow = g2.TopRow - 2 'move list top 2 rows up if not at start
       g2t = g2.TopRow
     Case Is = -120  'for each mousewheel movement backward (towards user) by 1 click
       If curow > 205 Then Exit Sub
       oldrow = g2.Row: newrow = oldrow + 2: curow = newrow 'move 2 rows down
       If oldrow > -1 Then HiRo oldrow, 0, 9
       HiRo newrow, 1, 9
       oldrow = newrow
       If newrow > g2.TopRow + 14 Then g2.TopRow = g2.TopRow + 2 'move list top 2 rows down if not at end
       g2t = g2.TopRow
     Case Else: Exit Sub
   End Select
 End If
End Sub

Private Sub g2_KeyDown(KeyCode%, Shf%)
 Dim i As Integer

 If g2.TextMatrix(g2.Row, 18) = "1" Then i = 1 Else i = 0 'get correction to row# of row pair
 g2.Row = g2.Row - i 'always force selected row of pair to be the 'thick' row
 Select Case KeyCode
   Case vbKeyF12 'View BOL
     g2pop3_Click
   Case vbKeyUp 'grid captures UpArrow key even before the form does (ie. Form_KeyDown defeated!!!) MS BUGG
   Case vbKeyInsert 'insert row
     g2pop1_Click
   Case vbKeyDelete 'remove row
     g2pop2_Click
   Case vbKeyDown 'can intercept DownArrow key before the grid does but not the UpArrow!!! MS BUGG
     KeyCode = 0: oldrow = g2.Row
     If oldrow > 205 Then  'at end of grid
       Exit Sub
     Else
       If g2.TopRow < 144 Then
         If g2.Row = g2.TopRow + 14 Then
           g2.TopRow = g2.TopRow + 2
         End If
       End If
     End If
     newrow = oldrow + 2: curow = newrow
     If oldrow > -1 Then HiRo oldrow, 0, 1
     HiRo newrow, 1, 1
     oldrow = newrow: g2t = g2.TopRow
   Case vbKeyPageUp, vbKeyPageDown 'same problem as Up/Down Arrow keys above!!  MS BUGG
     g2t = g2.TopRow
     If KeyCode = vbKeyPageUp And g2.Row < g2.TopRow Then 'overcome similar MS bug to Arrow-Up Key
       g2.TopRow = 0: g2t = 0
     End If
 End Select
End Sub

Private Sub g2_EnterCell()
 'a 'row' in this grid is actually 2 separate rows - a thick followed by a thin
 'always set row# to the 'thick' row of the pair - the 'corrected' row
 If g2.TextMatrix(g2.Row, 18) = "1" Then
   newrow = g2.Row - 1: g2.Row = newrow
 Else
   newrow = g2.Row
 End If
 If oldrow = newrow Then Exit Sub
 curow = newrow 'buff corrected current row
 If oldrow > -1 Then HiRo oldrow, 0, 0 'un-hilite old row
 HiRo newrow, 1, 0 'hilite new (current) row
 oldrow = newrow
End Sub

Private Sub HiRo(rw%, oo%, src%) 'rw=row; 00=leaving=0, entering=1; src=mouse=0,kboard=1
 Dim cl As Long
 Dim i As Integer, j As Integer, n As Integer, x As Integer
 'a 'row' in this grid is actually 2 separate rows - a thick followed by thin
 'hilite selected row
 g2pro.Visible = False
 i = rw: j = rw + 1
 If src = 0 Then curow = rw
 g2pro.Left = 210
 If oo = 0 Then 'leaving row 'rw' un-hilite
   For n = i To j
     g2.Row = n
     If g2.TextMatrix(n, 19) = "1" Then cl = &HDCE9EA Else cl = &HEFF4F5
     For x = 0 To 10
       g2.Col = x: g2.CellForeColor = &H10&: g2.CellBackColor = cl
     Next x
   Next n
 Else 'entering row 'rw' hilite
   For n = i To j
     g2.Row = n
     For x = 0 To 17
       g2.Col = x: g2.CellForeColor = &HFEFFFF: g2.CellBackColor = &H0
     Next x
   Next n
   g2.Row = i
   If g2.RowHeight(g2.TopRow) < 300 Then 'thin row (shouldn't happen)
     g2pro.Top = ((i - g2.TopRow) / 2 * 780) + 240 - 195 + frg1.Top
   Else 'thick row OK
     g2pro.Top = ((i - g2.TopRow) / 2 * 780) + 240 + frg1.Top 'place pro# entry textbox over grid
   End If
   curow = rw: prono = True: g2pro = g2.TextMatrix(i, 17): prono = False: g2pro.Visible = True
 End If
 g2t = g2.TopRow 'buff current toprow
End Sub

Private Sub g2_Scroll() 'only fires if scroll bar clicked, pgup/dn pressed or keyup/dn causes scrolling
 If g2.TopRow < g2t Then 'scroll up
   If g2.TextMatrix(g2.TopRow, 18) = "1" Then g2.TopRow = g2.TopRow - 1 'if toprow is thin, move down one to ensure full pair is visible (thick at top)
 ElseIf g2.TopRow > g2t Then 'scroll
   If g2.TextMatrix(g2.TopRow, 18) = "1" Then g2.TopRow = g2.TopRow + 1
 End If
 If curow >= g2.TopRow And curow <= g2.TopRow + 14 Then 'if hilited row in visible 8-rowpairs section of list
   g2pro.Top = ((curow - g2.TopRow) / 2 * 780) + 240 + frg1.Top 'place pro# entry textbox
   If g2.Row = curow Then
     g2pro.Left = 210: g2pro.Visible = True
   End If
 Else
   g2pro.Visible = False
 End If
 g2t = g2.TopRow 'buff current toprow
End Sub

Private Sub g2pro_KeyDown(KeyCode%, Shf%)
 Select Case KeyCode
   Case vbKeyF12: g2pop3_Click 'View BOL
   Case vbKeyDown, vbKeyReturn, vbKeyExecute
     If curow = 206 Then Exit Sub
     g2profoc = True
     newrow = curow + 2
     If oldrow > -1 Then HiRo oldrow, 0, 1
     HiRo newrow, 1, 1
     oldrow = newrow: curow = newrow
     If newrow - g2.TopRow > 15 Then
       g2.TopRow = g2.TopRow + 1: g2t = g2.TopRow
     End If
     DoEvents: DoEvents
     If g2.Visible Then g2pro.SetFocus
   Case vbKeyUp
     If curow = 0 Then Exit Sub
     g2profoc = True
     newrow = curow - 2
     If oldrow > -1 Then HiRo oldrow, 0, 1
     HiRo newrow, 1, 1
     oldrow = newrow: curow = newrow
     If curow < g2.TopRow Then
       g2.TopRow = g2.TopRow - 1: g2t = g2.TopRow
     End If
     DoEvents: DoEvents
     If g2.Visible Then g2pro.SetFocus
   Case vbKeyC
     If Shf = 2 Then
       Select Case Val(g2pro)
         Case 10000 To 9999999999#: Clipboard.Clear: Clipboard.SetText g2pro
       End Select
     End If
 End Select
End Sub
Private Sub g2pro_KeyPress(ka%)
 Select Case ka
   Case 8, 49 To 57, 65 To 90 'bkspace, 1-9, A-Z
   Case 48 'prevent leading zeros
     If g2pro.SelStart = 0 Then ka = 0
   Case 97 To 122: ka = ka - 32 'a-z convert to uppercase
   Case Else: ka = 0  'filter out all other chars
 End Select
End Sub

Private Sub g2pro_Change()
 Dim dm As String, xdte As String
 Dim k As Integer
   
 If Len(g2pro) < 5 Then
   If g2.TextMatrix(curow, 2) <> "" Then Clrg2 curow
   Exit Sub
 End If
 If prono Then Exit Sub 'recursion protection
 
 '####
 Select Case Val(g2pro)
   Case 1000000000 To 5999999999#
     rs.Open "select pronumber from probill where pronumber in ('" & g2pro & "','HMES" & g2pro & "')"
     If Not rs.EOF Then
       If Left$(rs(0), 4) = "HMES" Then
         dm = rs(0): rs.Close: g2pro = dm: Exit Sub
       End If
     End If
     rs.Close
 End Select
 '####
 
 'check if already entered
 For i = 0 To 206 Step 2
   If i <> curow Then
     If g2.TextMatrix(i, 17) = g2pro Then
       MsgBox "Pro# " & g2pro & " Already Entered at Row " & CStr((i + 2) / 2)
       Exit Sub
     End If
   End If
 Next i
 
 If Len(g2pro) = 16 Then g2pro.FontSize = 8 Else g2pro.FontSize = 10
 
 If g2.TextMatrix(curow, 17) <> "" Then
   Clrg2 curow
   WrTotWtCube
 End If
 
 'get probill record
 rs.Open "select * from probill where pronumber='" & g2pro & "'"
 If rs.EOF Then
   rs.Close: Exit Sub
 End If
 
 'check pro
 If Not loading And IsDate(rs(25)) = True And Len(g2pro) > 7 Then
   Select Case DateDiff("d", DateValue(rs(25)), DateValue(Now))
     Case Is > 365
       MsgBox "Pro is Older Than 1 Year. Please Check your Entry.", , "Manifest"
     Case Is > 185
       rsb.Open "select code from probdel where pro='" & g2pro & "' and code in ('*','D')"
       If Not rsb.EOF Then
         MsgBox "Pro is Older Than 6 Months and in Delivery Status. Please Check your Entry.", , "Manifest"
         rsb.Close: rs.Close: Exit Sub
       End If
       rsb.Close
   End Select
 End If
 
 Select Case Val(dest) 'check for EXLA, RDWY beyond pros on trips to EXLA-BUF, RDWY-BUF
   Case 54, 55
     If Len(rs(87)) >= 10 Then g2.TextMatrix(curow, 20) = rs(87)
 End Select
 
SKPG2PROCHK:
 
 DispProRow
 rs.Close
 If Not loading Then WrTotWtCube
 
End Sub

Private Sub WrTotWtCube()
 Dim j As Integer, cnt As Integer, cucnt As Integer
 Dim wt As Long, cube As Long
 wt = 0: cnt = 0: cube = 0: cucnt = 0
 For j = 0 To 206 Step 2
   If g2.TextMatrix(j, 5) <> "" Then
     cnt = cnt + 1: wt = wt + Val(Replace(g2.TextMatrix(j, 5), Chr(13), ""))
     If Val(g2.TextMatrix(j, 14)) > 0 Then
       cucnt = cucnt + 1: cube = cube + Val(g2.TextMatrix(j, 14))
     End If
   End If
 Next j
 twt = CStr(wt): tsh = CStr(cnt): tcu = CStr(cube): tcucnt = CStr(cucnt) & "/" & CStr(cnt)
End Sub

Private Function ChkProRow(pr$, rw%, typ%) As Boolean
 Dim dm As String, xdte As String, scod As String, pcod As String, pdte As String
 Dim j As Integer, k As Integer
 
 'check for void bill
 If InStr(g2.TextMatrix(rw, 2), "VOID PROBILL") > 0 Then
   MsgBox "VOID Probill " & pr & " on Row " & g2.TextMatrix(rw, 0) & " Cannot be Manifested!", , "Manifest"
   ChkProRow = True: Exit Function
 End If
 
 'get status from probill record (status history is already displayed for pro in grid)
 dm = "select funds,deldate from probill where pronumber='" & pr & "'"
 rsb.Open dm
 If rsb.EOF Then
   rsb.Close
   MsgBox "Internal Error: Probill Record Not Found for Pro: " & pr & " at Row " & g2.TextMatrix(rw, 0) & vbCrLf & vbCrLf & " ** Save Aborted **", , "Manifest"
   ChkProRow = True: Exit Function
 End If
 pcod = rsb(0)
 If IsNull(rsb(1)) = True Then pdte = "" Else pdte = Format$(DateValue(rsb(1)), "DD-MMM-YY")
 rsb.Close
 
 'get most-recent status history code from col#8 in grid
 If g2.TextMatrix(rw, 9) <> "" Then
   j = InStr(g2.TextMatrix(rw, 9), Chr(13))
   If j < 3 Then k = j + 1 Else k = 1
   scod = Mid$(g2.TextMatrix(rw, 9), k, 1)
 End If
  
 'check if used in another manifest
 rsb.Open "select trip,date,unit,trailer from trippro where pro='" & pr & "' order by pro, date desc limit 1"
 If Not rsb.EOF Then 'found existing record
   rsc.Open "select date_modified,hr_mod,min_mod,init from trip where assign='" & rsb(0) & "' order by assign, date desc limit 1"
   If rsc.EOF Then 'trip parent record Not Found! remove orphan trippro record
     dm = "delete from trippro where pro='" & pr & "'"
     cmd.CommandText = dm
     On Error Resume Next
     Err = 0
     cmd.Execute
     If Err <> 0 Then
       dm = "FAIL Remove Existing Un-Attached Pro Record!" & vbCrLf & vbCrLf & " ** Save Aborted ** Contact IT Support." & vbCrLf & vbrLf & Err.Description
       Err = 0: GoTo NOTMOV
     End If
     On Error GoTo 0
     rsc.Close: rsb.Close: GoTo SKPMFCHK
   End If
   
   'if here then valid records have been found of pro#'s previous manifesting and/or delivery
   If rsb(2) <> trip Then 'manifest to another trip#
     If IsNull(rsb(1)) = True Then xdte = "N/A" Else xdte = Format$(DateValue(rsb(1)), "DD-MMM-YY")
     dm = "Pro " & pr & " on Row " & g2.TextMatrix(rw, 0) & " is Already Manifested to: " & vbCrLf & vbCrLf & "  Unit "
     If rsb(2) = "" Then dm = dm & "N/A" Else dm = dm & rsb(2)
     dm = dm & "  Trailer "
     If rsb(3) = "" Then dm = dm & "N/A" Else dm = dm & rsb(3)
     dm = dm & "  Date " & xdte & vbCrLf & "  Last Modified "
     If IsNull(rsc(0)) = True Then xdte = "N/A" Else xdte = Format$(DateValue(rsc(0)), "DD-MMM-YY") & " " & Format$(Val(rsc(1)), "00") & ":" & Format$(Val(rsc(2)), "00")
     dm = dm & xdte & " by " & rsc(3) & vbCrLf & vbCrLf
     
     '21Nov13 - simply allow user to abort save
     rsc.Close: rsb.Close
     If typ = 0 Then dm = dm & "CLEAR PRO ?" Else dm = dm & "   ABORT SAVE ?"
     If MsgBox(dm, vbDefaultButton1 + vbExclamation + vbYesNo, "Pronumber Check") = vbYes Then ChkProRow = True Else ChkProRow = False
     Exit Function
     '21Nov13
     
     Select Case scod 'check current status & dates to see if pro can be moved from current trip to this trip
       Case "D", "L", "O", "Z" 'NEED GATE !! shows delivered, linehaul or out-on-delivery - so no auto move
         dm = dm & "  ** Save Aborted **": GoTo NOTMOV
       Case Else 'status history indicates pro could be OK to move
         Select Case pcod 'check probill record
           Case "D", "L", "O", "Z" 'NEED GATE!!! status from probill indicates move NOT OK
             dm = dm & "  ** Save Aborted **"
           Case Else 'probill last-status is OK
             dm = dm & "** Do You Wish to MOVE the Shipment to THIS Manifest? **"
             If MsgBox(dm, vbDefaultButton2 + vbQuestion + vbYesNo, "Move Shipment from Another Manifest") = vbNo Then GoTo NOTMOV
             dm = "delete from trippro where pro='" & pr & "'"
             cmd.CommandText = dm
             On Error Resume Next
             Err = 0
             cmd.Execute
             If Err <> 0 Then
               dm = "FAIL Remove Shipment from Existing Manifest!" & vbCrLf & vbCrLf & " ** Save Aborted ** Contact IT Support." & vbCrLf & vbrLf & Err.Description
               Err = 0: GoTo NOTMOV
             End If
             GoTo OKMOV
         End Select
     End Select
NOTMOV:
     rsc.Close: rsb.Close
     MsgBox dm, , "Manifest"
     ChkProRow = True: Exit Function
   End If
OKMOV:
   rsc.Close
 End If
 rsb.Close
 
SKPMFCHK:
End Function

Private Function GetDelTerm$(ac$, cy$, pv$, pz$, sp$)
 If Len(pz) < 3 Then GoTo TRYCY
 rsb.Open "select dterm from can_ctys_zone_fsa where fsa='" & Left$(pz, 3) & "'"
 If Not rsb.EOF Then
   Select Case rsb(0)
     Case "TR": GetDelTerm = "1"
     Case "MT": GetDelTerm = "12"
     Case "LN": GetDelTerm = "5"
     Case "WR": GetDelTerm = "6"
     Case "PI": GetDelTerm = "8"
     Case "AX": GetDelTerm = "10"
     Case "MI": GetDelTerm = "3"
     Case "ML": GetDelTerm = "4"
     Case "LA", "VA": GetDelTerm = "11"
     Case "QC": GetDelTerm = "14"
     Case "AC": GetDelTerm = "112"
     Case Else: rsb.Close: Pause 0.1: GoTo TRYCY
   End Select
 End If
 rsb.Close: Exit Function
TRYCY:
 rsb.Open "select dterm from can_ctys_zone_fsa where prov='" & pv & "' and city='" & mySav(cy) & "'"
 If Not rsb.EOF Then
   Select Case rsb(0)
     Case "TR": GetDelTerm = "1"
     Case "MT": GetDelTerm = "12"
     Case "LN": GetDelTerm = "5"
     Case "WR": GetDelTerm = "6"
     Case "PI": GetDelTerm = "8"
     Case "AX": GetDelTerm = "10"
     Case "MI": GetDelTerm = "3"
     Case "ML": GetDelTerm = "4"
     Case "LA", "VA": GetDelTerm = "11"
     Case "QC": GetDelTerm = "14"
     Case "AC": GetDelTerm = "112"
     Case Else: rsb.Close: Pause 0.1: GoTo TRYZONPV
   End Select
 End If
 rsb.Close: Exit Function
TRYZONPV:
 If Val(ac) > 1 Then
   Select Case pv
     Case "NB", "NF", "NL", "NS", "PE"
       If sp = "QC" Then GetDelTerm = "12" Else GetDelTerm = "1"
       rsb.Close: Exit Function
     Case "AB", "BC", "MB", "SK", "YT", "NT", "NU"
       GetDelTerm = "1": rsb.Close: Exit Function
   End Select
   rsb.Open "select zone from sc where account_num='" & ac & "'"
   If Not rsb.EOF Then
     Select Case rsb(0)
       Case "TORONTO", "BURLINGTON", "ST CATHARINES", "BARRIE", "PETERBOROUGH", "SUDBURY", "TIMMINS", "SAULT STE MARIE", "THUNDER BAY", "KENORA", "OWEN SOUND", "ORILLIA"
         GetDelTerm = "1"
       Case "MONTREAL", "MONTREAL SUBS", "DRUMMONDVILLE", "VALLEYFIELD", "BEAUCE", "GRANBY", "SHERBROOKE", "TROIS RIVIERES", "RIMOUSKI", "SAGUENAY", "MT LAURIER", "VAL DOR", "SEPT ILES"
         GetDelTerm = "12"
       Case "QUEBEC CITY": GetDelTerm = "14"
       Case "LONDON", "KITCHENER": GetDelTerm = "5"
       Case "WINDSOR": GetDelTerm = "6"
       Case "OTTAWA", "KINGSTON": GetDelTerm = "10"
     End Select
   End If
   rsb.Close
 End If
End Function

Private Sub DispProRow()
 Dim k As Integer, skd As Integer, icnt As Integer, dayix As Integer
 Dim dm As String, dm2 As String, ap As String, flgs As String, cod As String, ityp As String
 Dim sk As Variant, ct As Variant, ot As Variant
 Dim aset As Boolean, specdte As Boolean, newsched As Boolean
 
 'write probill to grid
 'col 1 pro#/date/collect
 g2.TextMatrix(curow, 17) = rs(1) 'buff pro in hidden column
 dm = rs(1) & Chr(13)
 If Not loading And Len(g2.TextMatrix(curow, 20)) = 11 Then dm = dm & g2.TextMatrix(curow, 20)
 
 Select Case Val(dest)
   ''27Sep16 remv Total Lgst NDTL  Case 50, 52, 63, 67, 95, 96 'CLRK (CDTT), NDTL, WARD, SEFL, FAST(95,96) **************
   Case 34 'KIDY - moved from below
     If loading Then
       g2.TextMatrix(curow, 20) = rsc(1)
       dm = dm & rsc(1)
     End If
   Case 50, 63, 67, 95, 96 'CLRK (CDTT), NDTL, WARD, SEFL, FAST(95,96) **************
     If loading Then
       g2.TextMatrix(curow, 20) = rsc(1) 'buff beyond pro in hidden column
       g2.TextMatrix(curow, 21) = rs(45) 'buff freight dimensions in hidden column (if any)
       dm = dm & rsc(1)
     End If
   Case 77, 79, 80, 87, 81 '', 89 'Guilbault Mtl, Que, Tor; M-O(81) 31Jan17 89 chg to AVERY-Mentor
     If loading Then
       g2.TextMatrix(curow, 20) = rsc(1) 'buff beyond pro in hidden column
       dm = dm & rsc(1)
     End If
   Case 82 'RSUT 'PCX Winnipeg
    If loading Then
      'no beyond pro or other rules currently
    End If
   Case 85, 86 'AMOT Moncton, Dartmouth
     If loading Then
       g2.TextMatrix(curow, 20) = rsc(1) 'buff AMOT bynd pro in hidden column
       g2.TextMatrix(curow, 21) = rs(45) 'buff freight dimensions in hidden column
       dm = dm & rsc(1)
     End If
   Case 65 ''27Sep16 remv ODFL Albany   , 66 'ODFL **********
     If loading Then
       If rs(62) <> "" Then 'field 'edi_unit'  = ODFL pro, buff in hidden col
         g2.TextMatrix(curow, 20) = rs(62)
       ElseIf Len(rs(27)) = 3 Then 'get ODFL pro from 'pickup' + speedy pro fields
         g2.TextMatrix(curow, 20) = rs(27) & Right$(rs(1), 8)
       End If
       dm = dm & g2.TextMatrix(curow, 20)
     End If
   Case Else
     If Not loading And Len(g2.TextMatrix(curow, 20)) = 11 Then
       'fall-thru
     Else
       If IsNull(rs(25)) = True Then dm = dm & "N/A !!" Else dm = dm & Format$(DateValue(rs(25)), "YYYY-MM-DD")
       dm = dm & Chr(13)
       If Val(rs(20)) > 0 Then
         dm = dm & "Collect"
       Else
         Select Case rs(24) 'field 14 - billedparty
           Case "P": dm = dm & "PrePaid"
           Case "C"
             If Val(rs(20)) > 0 Then dm = dm & "Collect"
           Case "T": dm = dm & "3rdParty"
         End Select
       End If
     End If
 End Select
 
 'field 19= codcharges, field 20= collect
 If Val(rs(20)) > 0 Then
   g2.TextMatrix(curow, 8) = "Collect" & Chr(13) & Format$(rs(20), "#.00")
 ElseIf Val(rs(19)) > 0 Then
   g2.TextMatrix(curow, 8) = "COD" & Chr(13) & Format$(rs(19), "#.00")
 End If
 g2.TextMatrix(curow, 1) = dm 'write pro/date/collect
 If InStr(rs(33), "** VOID **") > 0 Then
   g2.TextMatrix(curow, 2) = "* VOID PROBILL *" & Chr(13) & rs(3)
 Else
   'col 2 'shipper name - auto-wraps in cell
   g2.TextMatrix(curow, 2) = rs(3)
 End If
 'col3 'consigne name/address
 g2.TextMatrix(curow, 3) = rs(9) & Chr(13) & rs(10) & Chr(13) & rs(11) & "  " & rs(12) & "  " & rs(13)
 g2.TextMatrix(curow, 23) = rs(12) 'buff cons_prov 22Aug18(DNY)
 
 'delivery terminal                cons acc#  city    prov   postal   shipprov
 g2.TextMatrix(curow, 22) = GetDelTerm(rs(8), rs(11), rs(12), rs(13), rs(6))
  
 'col4 pcs/type
 'pcs counts for any of 3 types - cartons, skids, other (drums, pails, other)
 If rs(14) > 0 Then ct = rs(14)   'cartons
 If rs(16) > 0 Then ot = rs(16)   'drums treated as other
 If rs(17) > 0 Then ot = ot + rs(17) 'pails treated as other
 If rs(18) > 0 Then ot = ot + rs(18) 'other
 If rs(15) > 0 Then sk = rs(15) 'skids
 dm = ""
 Select Case Val(ct)
   Case 1 To 9999: dm = Space(4 - Len(ct)) & ct & "-C"
   Case 10000 To 99999: dm = Space(5 - Len(ct)) & ct & "-C"
   Case Is > 99999: dm = Space(6 - Len(ct)) & ct & "-C"
 End Select
 dm = dm & Chr(13)
 Select Case Val(sk)
   Case 1 To 9999: dm = dm & Space(4 - Len(sk)) & sk & "-S"
   Case 10000 To 99999: dm = dm & Space(5 - Len(sk)) & sk & "-S"
   Case Is > 99999: dm = dm & Space(6 - Len(sk)) & sk & "-S"
 End Select
 dm = dm & Chr(13)
 Select Case Val(ot)
   Case 1 To 9999: dm = dm & Space(4 - Len(ot)) & ot & "-O"
   Case 10000 To 99999: dm = dm & Space(5 - Len(ot)) & ot & "-O"
   Case Is > 99999: dm = dm & Space(6 - Len(ot)) & ot & "-O"
 End Select
 g2.TextMatrix(curow, 4) = dm
 'wt
 g2.TextMatrix(curow, 5) = Chr(13) & rs(30) & Chr(13)
 'cubewt 'asweight' field
 g2.TextMatrix(curow, 14) = rs(37)
 'hazmat indicator
 g2.TextMatrix(curow, 16) = ChkHazMat
 
 'last 3 status codes
 rsb.Open "select statuscode,statusseq,date_ent,hr_ent,min_ent from stathist where pro='" & rs(1) & _
          "' order by date_ent desc, cast(hr_ent as unsigned) desc, cast(min_ent as unsigned) desc limit 3"
 If Not rsb.EOF Then
   k = 0: dm = ""
   Do
     k = k + 1: ap = rsb(0) '& " "
     If ap = "D" Then
       If rsb(1) = "22" Then ap = ap & "-POD" Else ap = ap & "-PLM"
     End If
     If IsNull(rsb(2)) = True Then
       ap = ap & "-N/A"
     Else
       ap = ap & "-" & Format$(DateValue(rsb(2)), "DDMMMYY") & " " & Format$(Val(rsb(3)), "00") & ":" & Format$(Val(rsb(4)), "00")
     End If
     dm = dm & ap
     If k < 3 Then dm = dm & Chr(13)
     rsb.MoveNext
   Loop Until rsb.EOF
   If k = 1 Then dm = Chr(13) & dm
   g2.TextMatrix(curow, 9) = dm
 End If
 rsb.Close
 
 'check for BOL scan (std. or STPC)
 ityp = "1"
 dm = "select ndx from img_sfmaster where pro_date='" & g2.TextMatrix(curow, 17) & _
      "' and ndx like 'SBOL%' and accnt like 'BOL-%' and bkup <> '9' and pgno='1' order by ndx desc limit 1"
 ir.Open dm
 If ir.EOF Then 'std. BOL not found
   Select Case Left$(g2.TextMatrix(curow, 17), 4)
     Case "EXLA", "HMES", "PYLE", "SEFL", "WARD": GoTo CHKUSPRTNRBOL
   End Select
   Select Case Val(g2.TextMatrix(curow, 17)) 'check pro# range
         'NewPenn/Avery, Holland & ODFL ranges
         ''Case 50000000# To 89999999#, 1000000000# To 4999999999#, 9000000000# To 9099999999#, 9300000000# To 9399999999#, 9900000000# To 9999999999#, 9210000000# To 9210099999#
     Case 1000000000# To 4999999999#
CHKUSPRTNRBOL:
       ir.Close: ityp = "2" 'CHECK STPC (northbound)
       dm = "select ndx from img_stpc where pro_date='" & g2.TextMatrix(curow, 17) & _
            "' and pgno='1' and inv_unit='BOL' and accnt <> 'VOID' order by ndx desc limit 1"
       ir.Open dm
       If Not ir.EOF Then 'STPC BOL image found
         g2.TextMatrix(curow, 10) = "*": g2.TextMatrix(curow, 11) = ir(0): g2.TextMatrix(curow, 15) = ityp
       Else
         ir.Close: ityp = "1" 'CHECK PAPS (southbound)
         
         dm = "select ndx from img_sfmaster where pro_date in ('" & g2.TextMatrix(curow, 17) & "','RDWY" & Mid(g2.TextMatrix(curow, 17), 5) & _
              "') and ndx like 'PAPSMF%' and accnt='PAPS BOL' and inv_unit='ORIG' and pgno='1' order by ndx desc limit 1"
              
              
         ir.Open dm
         If Not ir.EOF Then
           g2.TextMatrix(curow, 10) = "*": g2.TextMatrix(curow, 11) = ir(0): g2.TextMatrix(curow, 15) = ityp
         End If
       End If
     Case Else
   End Select
 Else
   g2.TextMatrix(curow, 10) = "*": g2.TextMatrix(curow, 11) = ir(0): g2.TextMatrix(curow, 15) = ityp
 End If
 ir.Close
         
 ''{@} get most-recent row/location record:  Date Time Term-Row/Trailer
 '                    0         1     2    3       4
 rsb.Open "select date_time,terminal,row,trailer,inits from dock_check where pro='" & _
           g2.TextMatrix(curow, 17) & "' order by pro, date_time desc limit 1"
 If rsb.EOF Then
   dm = "  N/A"
 Else
   If Not IsNull(rsb(0)) Then
     i = InStr(rsb(0), " ")
     If i > 0 Then
       dm = Format$(DateValue(Left$(rsb(0), i - 1)), "DDMMM") & " " & _
            Format$(TimeValue(Mid$(rsb(0), i + 1)), "Hh:Nn")
     Else
       dm = "??????? ??:??"
     End If
   Else
     dm = "??????? ??:??"
   End If
   Select Case Left$(rsb(1), 1) 'ensure correct 2-char terminal codes
     Case "A": dm2 = "BRO" 'dm2 = "AX" '909#
     Case "B": dm2 = "BRO"
     Case "L": dm2 = "LON" 'dm2 = "LN"
     Case "M": dm2 = "MTL" 'dm2 = "MT"
     Case "P": dm2 = "PIC"
     Case "T": dm2 = "TOR" 'dm2 = "TR"
     Case "V": dm2 = "LAC" '"VAU" in 2022
     Case "W": dm2 = "WIN" 'dm2 = "WR"
     Case "S": dm2 = "MIS"
     Case "I": dm2 = "MIL"
     Case Else: dm2 = "??"
   End Select
   dm = "  " & dm & " " & dm2 & "?"
   If Trim$(rsb(2)) <> "" Then
     dm = dm & "R? & Trim$(rsb(2))"
   ElseIf Trim$(rsb(3)) <> "" Then
     dm = dm & "T? & rsb(3)"
   Else
     dm = dm & "N/A"
   End If
   If rsb(4) <> "" Then dm = dm & " " & rsb(4) 'inits
 End If
 rsb.Close
 For k = 0 To 2: g2.TextMatrix(curow + 1, k) = dm: Next k 'write last location to merged cells
       
 'hrs rcvg - get from ops account-code records (after hazmat, etc. checks below)
 aset = False: dm = "": ap = "": cod = "": flgs = "*"
 
 'chk hazmat
 If Val(g2.TextMatrix(curow, 16)) > 0 Then flgs = flgs & "* HAZMAT *"
 
 'chk for heat/freezble
 If rsb.State <> 0 Then rsb.Close
 rsb.Open "select note from probnotes where pro='" & g2.TextMatrix(curow, 17) & "'"
 If Not rsb.EOF Then
   dm2 = Replace(rsb(0), vbCrLf, " ")
    '' see accnt codes
    ''If InStr(dm2, "TAILGATE REQ") > 0 Then flgs = flgs & "* TAILGATE REQD *"
   If InStr(dm2, "HEAT REQ") > 0 Or InStr(dm2, "HEAT HEAT") > 0 Then flgs = flgs & "* HEAT REQD *"
   If InStr(dm2, "FREEZABLE") > 0 Or InStr(dm2, "FREEZING") > 0 Then flgs = flgs & "* FREEZABLE *"
     
 Else
   For k = 43 To 45
     If InStr(rs(k), "HEAT") > 0 Then flgs = flgs & "* HEAT REQD *"
     If InStr(dm2, "FREEZABLE") > 0 Or InStr(dm2, "FREEZING") > 0 Then flgs = flgs & "* FREEZABLE *"
   Next k
 End If
 rsb.Close
 
 'chk for guaranteed
 If InStr(rs(33), "GUARAN") > 0 Then
   flgs = flgs & "* " & Replace(rs(33), "  ", " ") & " *"
 End If
 
 'check CTC - Call in Waiting Time notice 19Sep16(DNY)
 If Left$(Replace(Left$(g2.TextMatrix(curow, 3), 4), "CTC#", "CTC "), 4) = "CTC " Or Left$(g2.TextMatrix(curow, 3), 13) = "CANADIAN TIRE" Then
   flgs = flgs & "* Call-In ALL Waiting Time *"
 End If
 
 'check consignee account-codes, shipment special requirements, etc. - build shipment special handling flags
 If Val(rs(8)) = 0 Then
   ''If flgs = "*" Then
   GoTo SKPOLDOPSCOD 'cons accnt# must be non-zero for further consignee processing
 End If
 
 dm2 = Format$(mdte, "YYYY-MM-DD"): specdte = False
 dm = "select * from scops2date where date='" & dm2 & "' and accno='" & rs(8) & "'"
 rsb.Open dm
 If Not rsb.EOF Then
   specdte = True 'set indicator
   If rsb(5) = 1 Then 'closed
     cod = "SITE CLSD" & Chr(13) & dm2: flgs = flgs & "* SITE CLOSED " & dm2 & " *"
   Else 'open, check for rcvg hrs and/or appointment
     'check/write receiving hrs
     If Val(rsb(6)) > 0 And Val(rsb(7)) > 0 Then
       g2.TextMatrix(curow, 7) = Chr(13) & rsb(6) & "-" & rsb(7)
       If rsb(10) <> 1 Then flgs = flgs & "* SPECIAL RCVG HRS *"
     End If
     If rsb(10) = 1 Then 'appt req'd
       If rsb(12) = 1 Then 'appt only if not in rcvg hrs
         cod = "APPT IFNOT" & Chr(13) & "RCVG HRS"
         flgs = flgs & "* APPT REQD IF NOT RCVG HRS *"
       Else 'unconditional
         cod = "APPT REQD"
         flgs = flgs & "* APPT REQD *"
       End If
     End If
   End If
   rsb.Close: GoTo SKPOLDOPSCOD
 End If
 rsb.Close
 
 If Not specdte Then 'no specific override set previously for this manifest date
    'check if new site-access/appointment regular config exists
   '             0   1   2   3   4
   dm = "select mon,tue,wed,thu,fri," 'site open
   '             5          6          7          8          9
   dm = dm & "mon_del_op,tue_del_op,wed_del_op,thu_del_op,fri_del_op," 'rcvng hrs open
   '             10         11         12         13         14
   dm = dm & "mon_del_cl,tue_del_cl,wed_del_cl,thu_del_cl,fri_del_cl," 'rcvng hrs clsd
   '             15          16          17          18          19
   dm = dm & "mon_delappt,tue_delappt,wed_delappt,thu_delappt,fri_delappt," 'appt reqd
   '             20        21        22        23        24
   dm = dm & "mon_delsk,tue_delsk,wed_delsk,thu_delsk,fri_delsk," 'appt reqd if skids > value
   '             25         26         27         28         29   'appt reqd if not in rcvng hrs
   dm = dm & "mon_deltim,tue_deltim,wed_deltim,thu_deltim,fri_deltim from scops2 where accno='" & rs(8) & "'"
   newsched = False: rsb.Open dm
   If Not rsb.EOF Then
     newsched = True
     dayix = Weekday(mdte, vbMonday) - 1 '0=Monday, etc.
     If dayix < 5 Then 'not weekend
       '1. check if site open
       If Val(rsb(dayix)) = 1 Then 'site closed
         cod = "SITE CLSD": flgs = flgs & "* SITE CLOSED *" 'set appt column entry & account code flag text
       Else 'site open
         'check rcvg hrs
         If (Val(rsb(dayix + 5)) + Val(rsb(dayix + 10))) > 0 Then 'hrs entered
           g2.TextMatrix(curow, 7) = Chr(13) & rsb(dayix + 5) & "-" & rsb(dayix + 10) 'write to manifest column
           ''''If rsb(dayix + 15) <> 1 Then flgs = flgs & "* SPECIAL RCVG HRS *" 'if not appt then flag simply as rcvng hrs
         End If
         'check appt
         cod = ""
         If Val(rsb(dayix + 15)) = 1 Then 'appt reqd
           If Val(rsb(dayix + 25)) = 1 Then 'appt conditional on being outside rcvng hrs
             cod = "APPT IFNOT" & Chr(13) & "RCVG HRS"
             flgs = flgs & "* APPT REQD IF NOT RCVG HRS *"
           End If
           If Val(rsb(dayix + 20)) > 0 Then 'conditional on skid count > entry value
             If cod = "" Then cod = "APPT IF >" & Chr(13) & rsb(dayix + 20) & " SK"
             flgs = flgs & "* APPT REQD IF SK > " & rsb(dayix + 20) & " *"
           End If
           If Val(rsb(dayix + 25)) = 0 And Val(rsb(dayix + 20)) = 0 Then 'appt always reqd
             cod = "APPT REQD": flgs = flgs & "* APPT REQD *"
           End If
         End If
       End If
     End If
     rsb.Close: GoTo SKPOLDOPSCOD
   End If
   rsb.Close
 End If
  
 'check for old site-access/appointment config
 If Val(rs(2)) > 0 Then
   flgs = flgs & GetOpsNonCompFlags
   dm = "select * from scops where accno='" & rs(8) & "'"
   rsb.Open dm
   If rsb.EOF Then
     rsb.Close
     flgs = flgs & GetOpsShipFlags
     GoTo SKPACCOD
   End If
 End If
 If rsb.State = 0 Then GoTo SKPOLDOPSCOD
 If rsb(14) = 1 Then 'appt code set
   cod = "" 'build appt info from code (may not apply if overriden by shipment BOL)
   If rsb(15) = 1 Then 'check pcs cnt constraint
     If Val(rsb(16)) > 0 Then 'pcs count entered
       If Val(rsb(17)) > 0 Then 'use skid count, check if skid count from bill conforms to constraint
         If sk > rsb(16) Then 'appt constraint met
           cod = "** PAL > " & rsb(16) & "**" 'build text as warning
         Else 'no appt according to skid count
          cod = "OK PAL < " & CStr(rsb(16) + 1) 'build OK string
         End If
       End If
       If Val(rsb(18)) > 0 Then 'do same as above for pcs count
         If (ct + ot) > rsb(16) Then
           cod = "** PCS > " & rsb(16) & "**"
         Else
           cod = "OK PCS < " & CStr(rsb(16) + 1)
         End If
       End If
     End If
   End If
   If rsb(19) = 1 Then 'check time constraint
     If rsb(29) = 1 Then 'check if contrained to specific days
       If Val(rsb(30)) + Val(rsb(31)) + Val(rsb(32)) + Val(rsb(33)) + Val(rsb(34)) > 0 Then 'day-constrained
         If cod <> "" Then cod = cod & Chr(13) & "NOT"  'move to next line if pcs count constraint found above
         If Val(rsb(30)) = 1 Then cod = cod & " Mon"
         If Val(rsb(31)) = 1 Then cod = cod & " Tue"
         If Val(rsb(32)) = 1 Then cod = cod & " Wed"
         If Val(rsb(33)) = 1 Then cod = cod & " Thu"
         If Val(rsb(34)) = 1 Then cod = cod & " Fri"
       Else 'time-constrained
         If cod <> "" Then cod = cod & Chr(13) & "OK "
         If rsb(20) <> "" Then
           If Len(rsb(20)) = 3 Then cod = cod & "0" & rsb(20) Else cod = cod & rsb(20)
         End If
         If rsb(21) <> "" Then
           cod = cod & "-"
           If Len(rsb(21)) = 3 Then cod = cod & "0" & rsb(21) Else cod = cod & rsb(21)
         End If
       End If
     Else 'times are open/close times that are not specifically appointment-related
       If rsb(20) <> "" And rsb(21) <> "" Then 'use as rcvng hrs
         dm = Chr(13)
         If Len(rsb(20)) = 3 Then dm = "0" & rsb(20) Else dm = rsb(20)
         dm = dm & "-"
         If Len(rsb(21)) = 3 Then dm = dm & "0" & rsb(21) Else dm = dm & rsb(21)
         g2.TextMatrix(curow, 7) = dm
       End If
     End If
   End If
   If cod = "" Then cod = "APPT REQD"
 End If
 If rsb.State <> 0 Then rsb.Close
 
SKPOLDOPSCOD:

 'check special accounts:
 '1. Nippon Express
 If InStr(g2.TextMatrix(curow, 2), "NIPPON EXP") > 0 Then
   flgs = flgs & "* DO NOT STACK *"
 End If

 'check shipper straight/trailer size flags
 If Val(rs(2)) > 0 Then flgs = flgs & GetOpsShipFlags
 'check consignee delivery-restraint & non-compliance flags
 If Val(rs(8)) > 0 Or Val(rs(2)) > 0 Then flgs = flgs & GetOpsNonCompFlags
  
 'skip here if account# zero or not found
SKPACCOD:

 'consolidate & display ops flags for shipment
 If flgs = "*" Then flgs = " " Else flgs = flgs & "*" 'NOTE: need at least 1 char in cell to merge into 1
 dm = ChkSpecReq ' check if shipment ID'd as 'Special Requirements'
 If flgs = " " Then
   If dm <> "" Then flgs = dm
 Else
   If dm <> "" Then
     If Len(dm) + Len(flgs) > 90 Then flgs = dm Else flgs = dm & " " & flgs 'spec reqs trump std. flags
   End If
 End If
 ''For k = 1 To 9
 For k = 3 To 9  '{@} was 1 to 9 write identical string into each 'merge' cell to trigger the merger effect across them
   g2.TextMatrix(curow + 1, k) = flgs
 Next k
  
 If specdte Or newsched Then
   If cod <> "" Then
     If InStr(cod, "APPT") Then g2.TextMatrix(curow, 13) = "1" 'mark bill as appt
   End If
   If newsched Then
     If Not IsNull(rs(41)) Then
       If rs(41) <> "" Then GoTo ACTAPPT
     End If
   End If
   If cod <> "" Then
     '''If InStr(cod, "APPT") Then g2.TextMatrix(curow, 13) = "1" 'mark bill as appt
     g2.TextMatrix(curow, 6) = Chr(13) & cod
   End If
   If rsb.State <> 0 Then rsb.Close
   Exit Sub
 End If
 
ACTAPPT:
 'check actual appt info as entered in bill
 dm = "": ap = "" 'build appt info text for actual appt from bill
 'NOTE: ops record left open
 If Not IsNull(rs(41)) Then 'appt_date
   If rs(41) <> "" Then 'appt date entered
     aset = True 'indicate appt set by bill
     If DateValue(rs(41)) <> mdte Then 'hilite date mismatch between appt & manifest date IF not linehaul
       Select Case Val(orig)
         Case 0 'nothing
         Case 1 'TOR
           Select Case Val(dest)
             Case 0 'nothing
             Case 1, 71, 77 'local trip
               ap = "*Chk Date*"
             Case Else 'linehaul
           End Select
         Case Else
           Select Case Val(dest)
             Case 0 'nothing
             Case Val(dest) 'local trip
               ap = "*Chk Date*"
             Case Else 'linehaul
           End Select
       End Select
     End If
     If Val(rs(40)) > 0 Then
       If Len(rs(32)) = 3 Then dm = "0" & rs(32) Else dm = rs(32)
       If Len(rs(40)) = 3 Then dm2 = "0" & rs(40) Else dm2 = rs(40) 'apptmnt_time' field (smallint-4)
       If Val(dm) > 500 And Val(dm) < 1000 Then 'look for appt time between 5:01AM and 9:59AM
         If Val(dm2) > Val(dm) And Val(dm2) < 1000 Then
           g2.TextMatrix(curow, 12) = "1"
         End If
       End If
       dm = dm & "-" & dm2
     Else 'From-time-only indicates exact time or numeric code (eg. 1 = ASAP)
       If Len(rs(32)) = 3 Then dm = "0" & rs(32) Else dm = rs(32)
       If Val(dm) > 500 And Val(dm) < 1000 Then g2.TextMatrix(curow, 12) = "1" 'mark all appointment deliverys between 5:01AM & 9:59AM
     End If
     If dm = "" Then dm = "** N/A **" 'appt time(s)/code not entered in bill!
     If ap = "" Then 'date is OK from test above
       ap = Format$(DateValue(rs(41)), "YYYY-MM-DD") & Chr(13) & dm & Chr(13) & Trim$(rs(42))
     Else 'date mismatch - include warning in text
       ap = Format$(DateValue(rs(41)), "YYYY-MM-DD") & Chr(13) & dm & Chr(13) & ap
     End If
   End If
 End If
 If aset Then
   g2.TextMatrix(curow, 6) = ap 'write appt info
   g2.TextMatrix(curow, 13) = "1" 'mark bill as appt
 Else
   If cod <> "" Then g2.TextMatrix(curow, 6) = "Not Set" & Chr(13) & cod
 End If
 
End Sub

Private Function GetOpsShipFlags$()
 rsb.Open "select cons_trlr from scops where accno='" & rs(2) & "'"
 If Not rsb.EOF Then
   If Val(rsb(0)) > 0 Then GetOpsShipFlags = "* PKUP: " & rsb(0) & "' MAX Trailer *"
 End If
 rsb.Close
End Function

Private Function GetOpsNonCompFlags$()
 Dim dm As String, dm2 As String

 'check for shipper-driven special instructions flag
 rsb.Open "select msgasship1 from scops where accno='" & rs(2) & "' and msgasship='1' and msgship1asremark='1'"
 If Not rsb.EOF Then dm2 = "* " & rsb(0) & " *"
 rsb.Close
 If rs(8) = 0 Then GoTo WRSHPMSG
 'check consignee flags
 '                0         1        2         3         4         5         6       7       8         9             10              11            12
 dm = "select cons_tail,cons_str,cons_cube,cons_balm,cons_roll,cons_trlr,cons_am,cons_pm,cons_pump,cons_sideways,cons_nodblstk,msgcons1asremark,msgascons1 from scops where accno='" & rs(8) & "'"
 rsb.Open dm

 dm = ""
 If Not rsb.EOF Then
   If rsb(0) = 1 Then dm = dm & "* TAILGATE REQ'D *"
   If rsb(1) = 1 Then dm = dm & "* STRAIGHT TRUCK REQ'D *"
   If rsb(2) = 1 Then dm = dm & "* CUBE VAN REQ'D *"
   If rsb(3) = 1 Then dm = dm & "* HAND BALM REQ'D *"
   If rsb(4) = 1 Then dm = dm & "* ROLL-UP REQ'D *"
   If Val(rsb(5)) > 0 Then dm = dm & "* " & rsb(5) & "' MAX Trailer *"
   If rsb(6) = 1 Then dm = dm & "* AM DELIVERY *"
   If rsb(7) = 1 Then dm = dm & "* PM DELIVERY *"
   If rsb(8) = 1 Then dm = dm & "* PUMP TRUCK REQ'D *"
   If rsb(9) = 1 Then dm = dm & "* DO NOT LOAD SIDEWAYS *"
   If rsb(10) = 1 Then dm = dm & "* DO NOT DOUBLE-STACK *"
   If dm2 <> "" Then dm = dm & dm2 'write shipper-based instruction here
   If rsb(11) = 1 Then dm = dm & "* " & rsb(12) & " *"  'Line 1 Delivery Instruction Accnt Code on Ops1 tab
 Else
WRSHPMSG:
   If dm2 <> "" Then dm = dm & dm2 'write shipper-based instruction
 End If
 If rsb.State <> 0 Then rsb.Close
 rsb.Open "select accno from scnoncomp where accno='" & rs(8) & "'"
 If Not rsb.EOF Then dm = dm & "* NON-COMPLIANCE *"
 rsb.Close
 GetOpsNonCompFlags = dm
End Function

Private Function ChkSpecReq$()
 rsb.Open "select sr1, sr2, sr3 from specreqs where pro='" & g2pro & "'"
 If rsb.EOF Then
   rsb.Close: Exit Function
 End If
 ChkSpecReq = rsb(0) & " " & rsb(1) & " " & rsb(2)
 rsb.Close
End Function

Private Function ChkHazMat$()
 'look for any indication that shipment is Hazmat
 If Val(rs(57)) > 0 Or rs(57) = "Y" Then GoTo CHM 'Hazmat Y/N probill entry, 'laststatus' field
 rsb.Open "select note from probnotes where pro='" & g2pro & "'"
 If Not rsb.EOF Then
   If HazTest(Replace(rsb(0), vbCrLf, " ")) Then
     rsb.Close: GoTo CHM
   End If
 End If
 rsb.Close

 If HazTest(rs(33)) Then GoTo CHM   'remark'
 If HazTest(rs(21)) Then GoTo CHM   'desc1'
 Exit Function
CHM:
 ChkHazMat = "1"
End Function

Private Function HazTest(fx$) As Boolean
 If InStr(fx, "DANGEROUS GOODS") > 0 Then
   HazTest = True
 ElseIf InStr(fx, "HAZARDOUS") > 0 Then
   If InStr(fx, "NON-HAZ") = 0 Then HazTest = True
 End If
End Function

Private Sub dest_Change()
 Dim dm As String

 If Val(dest) = 11 Then
   l_de = "LAC-Local"
   If Val(orig) > 0 Then
     If Val(orig) <> 11 Then
       orig = "11": Exit Sub
     End If
   End If
 ElseIf Val(orig) = 11 And dest <> "" Then
   dest = "11": Exit Sub
 Else
   l_de = DispLH(dest)
 End If
   ''l_de = DispLH(dest)
 If l_de = "" Then dest.ForeColor = &H80 Else dest.ForeColor = &H0
 DispTripType
End Sub

Private Sub orig_Change()
 If Val(orig) = 14 Then
   orig = "": Exit Sub
 End If
 If Val(orig) = 11 Then
   l_or = "LAC-Local"
   If Val(dest) > 0 Then
     If Val(dest) <> 11 Then
       dest = "11": Exit Sub
     End If
   End If
 ElseIf Val(dest) = 11 And orig <> "" Then
   orig = "11": Exit Sub
 Else
   l_or = DispLH(orig)
 End If
   ''l_or = DispLH(orig)
 If l_or = "" Then orig.ForeColor = &H80 Else orig.ForeColor = &H0
 DispTripType 'display trip description from orig/dest codes
End Sub

Private Sub DispTripType()
 Dim cl As Long
 If orig = "" Or dest = "" Or l_or = "" Or l_de = "" Then
   l_od = "": Exit Sub
 End If
 
 '19May20(DNY) allow LH moves in USA between US Partner terminals
 Select Case Val(orig)
   Case 45, 46, 47, 51, 52, 53, 54, 55, 57 To 61, 63, 67 '19May2020 -> PYLE 45 to 47, EXLA 52 to 54, YRC 55, HMES 57 to 61, WARD 63, SEFL 67
     Select Case Val(dest)
       Case 45, 46, 47, 51, 52, 53, 54, 55, 57 To 61, 63, 67
         l_od = "Inter-USA LINEHAUL TRIP -": cl = &H0: GoTo DTTCL
     End Select
 End Select

 Select Case Val(orig)
   Case 31
     If Val(dest) = 0 Then dest = "31"
     l_od = "* STORAGE TRAILER *": cl = &H0: GoTo DTTCL
   Case 44
      If Val(dest) = 0 Then dest = "44"
      l_od = "* LOCAL STRIP *": cl = &H0: GoTo DTTCL
   Case 37, 38, 69, 71, 72, 73
     If Val(orig) = Val(dest) Then 'valid cust pickup at dock codes
       l_od = "CUST DOCK PKUP -": cl = &H603020: GoTo DTTCL
     Else
       Select Case Val(dest)
         Case 45, 46, 47, 76, 77, 78, 79, 80, 81, 82, 85, 86, 94, 95, 96, 99 ' valid transfer to agent/partner codes
           l_od = "AGNT TRNSFR AT DOCK -": cl = &H603020: GoTo DTTCL
         Case Else 'all other combos with orig=dock are invalid
           l_od = "* INVALID ORIG -> DEST * -": cl = &H202AC0: GoTo DTTCL
       End Select
     End If
   Case 68
     If Val(dest) = 68 Then
       l_od = "SAMEDAY -": cl = &H603020: GoTo DTTCL
     Else
       l_od = "* INVALID ORIG -> DEST * -": cl = &H202AC0: GoTo DTTCL
     End If
   Case 77, 78
     l_od = "* INVALID ORIG * -": cl = &H202AC0: GoTo DTTCL
   Case 1 'Tor terminal code
     Select Case Val(dest)
       Case 37, 38, 68, 69, 71, 72, 73
         l_od = "* INVALID ORIG -> DEST * -": cl = &H202AC0: GoTo DTTCL
       Case 1, 77
         l_od = "LOCAL TRIP -": cl = &H0: GoTo DTTCL
       Case Else
         l_od = "LINEHAUL TRIP -": cl = &H0: GoTo DTTCL
     End Select
   Case 3, 4, 5, 6, 8, 10, 11, 12
     If Val(orig) = Val(dest) Then
       l_od = "LOCAL TRIP -": cl = &H0: GoTo DTTCL
     Else
       If Val(orig) = 11 Or Val(dest) = 11 Then
         l_od = "* INVALID ORIG -> DEST * -": cl = &H202AC0: GoTo DTTCL
       End If
       Select Case Val(dest)
         Case 37, 38, 68, 69, 71, 72, 73
           l_od = "* INVALID ORIG -> DEST * -": cl = &H202AC0: GoTo DTTCL
         Case Else
           l_od = "LINEHAUL TRIP -": cl = &H0: GoTo DTTCL
       End Select
     End If
     Exit Sub
   Case 99 'Misc
     If Val(dest) = 99 Then '26May2020 Misc-Misc trip
       l_od = "MISC-MISC TRIP": cl = &H0: GoTo DTTCL
     Else
       l_od = "MISC TRIP": cl = &H0: GoTo DTTCL  '10Jun20 was  l_od = "* UNKNOWN TRIP TYPE * -": cl = &H202AC0: GoTo DTTCL
     End If
   Case Else 'all other origin codes
     Select Case Val(dest)
       Case 37, 38, 71, 72, 73 'only cust pickups have dock pickup codes as destinations
         l_od = "* INVALID ORIG -> DEST * -": cl = &H202AC0: GoTo DTTCL
       Case 1, 3, 4, 5, 6, 8, 10, 11, 12 'Speedy terminal codes
         If Val(orig) = Val(dest) Then
           l_od = "LOCAL TRIP -"
         Else
           l_od = "LINEHAUL TRIP -"
         End If
         cl = &H0: GoTo DTTCL
       Case Else
         l_od = "* UNKNOWN TRIP TYPE * -": cl = &H202AC0: GoTo DTTCL
     End Select
 End Select
DTTCL:
 l_od.ForeColor = cl&
End Sub

Private Function DispLH$(x As Variant)
 Dim j As Integer
 If x = "" Then Exit Function
 For j = 0 To lst_or.ListCount - 1
   If lst_or.ItemData(j) = x Then
     DispLH = Trim$(Left$(lst_or.List(j), 7)): Exit Function
   End If
 Next j
End Function

Private Sub t_fpro_Change()
 If Len(t_fpro) < 5 Then
   t_fpro.ForeColor = &H80: Exit Sub
 End If
 rs.Open "select pronumber from probill where pronumber='" & t_fpro & "'"
 If rs.EOF Then t_fpro.ForeColor = &H80 Else t_fpro.ForeColor = &H0
 rs.Close
End Sub
Private Sub t_fpro_GotFocus()
  t_fpro.BackColor = &HFFFFFF
End Sub
Private Sub t_fpro_LostFocus() 'test trailer no.
 t_fpro.BackColor = &HD6E6E6
End Sub
Private Sub t_fpro_KeyDown(KeyCode%, Shf%)
 Select Case KeyCode
   Case vbKeyEscape: t_fpro = ""
   Case vbKeyReturn, vbKeyExecute: c_fgo_Click
   Case vbKeyV
     If Shf = 2 And Len(Clipboard.GetText) > 6 Then t_fpro = Clipboard.GetText
 End Select
End Sub
Private Sub t_fpro_KeyPress(ka%)
 Select Case ka
   Case 8, 48 To 57, 65 To 90 'bkspace, 0-9, A-Z
   Case 97 To 122: ka = ka - 32 'a-z convert to uppercase
   Case Else: ka = 0  'filter out all other chars
 End Select
End Sub
Private Sub t_tom_KeyPress(KeyAscii%)
 Select Case KeyAscii
   Case 8, 48 To 57
   Case Else: KeyAscii = 0
 End Select
End Sub
Private Sub t_tom_LostFocus()
 If t_tom = "" Then t_tom = "0"
End Sub

Private Sub t_unfr_Change()
 Dim j As Integer
 If notunfr Then Exit Sub
 If t_unfr = "" Then
   lst_un.ListIndex = 0
 Else
   For j = 1 To lst_un.ListCount - 1
     If t_unfr = lst_un.List(j) Then
     'If Val(t_unfr) = Val(Left$(lst_un.List(j), 6)) Then
       lst_un.ListIndex = j: Exit For
     End If
   Next j
 End If
End Sub
Private Sub t_unfr_GotFocus()
 t_unfr.SelStart = Len(t_unfr)
End Sub
Private Sub t_unfr_KeyDown(KeyCode%, Shf%)
 Select Case KeyCode
   Case vbKeyEscape: c_cancun_Click
   Case vbKeyReturn, vbKeyExecute
     If lst_un.ListIndex = 0 Then
       If Val(t_unfr) = "" Then lst_un_DblClick
     Else
       If Val(Left$(lst_un.Text, 6)) = Val(t_unfr) Then lst_un_DblClick
     End If
 End Select
End Sub
Private Sub t_unfr_KeyPress(KeyAscii%)
 KeyAscii = NumFilt(KeyAscii, True, t_unfr.SelStart)
End Sub

Private Sub lst_un_Click()
 notunfr = True
 If lst_un.ListIndex = 0 Then
   t_unfr = ""
 Else
   Select Case Left$(Trim$(lst_un.Text), 3)
     Case "TOR", "MTL", "LON", "PIC", "WIN", "BRO": t_unfr = Trim$(Left$(lst_un.Text, 8))
     Case Else:   t_unfr = Trim$(Left$(lst_un.Text, 6))
   End Select
 End If
 notunfr = False
End Sub
Private Sub lst_un_DblClick()
 Dim dm As String, dm2 As String, uni As String
 Dim n As Integer
 Dim s() As String
 
 If InStr(lst_un.Text, "- UN") > 0 Then
   unit = Trim$(lst_un.Text): nottun = True: t_un = unit: nottun = False
 Else
   uni = Replace(Trim$(lst_un.Text), "  ", " ")
   s = Split(uni, " "): uni = s(0)
   Select Case Left$(uni, 3)
     Case "TOR", "MTL", "LON", "PIC", "BRO", "WIN"
       Select Case Mid$(uni, 4, 1)
         Case 1: nottun = True: t_un = "SPDY " & Left$(uni, 3) & " LOC ST  " & uni: nottun = False    'Straight
         Case 2: nottun = True: t_un = "SPDY " & Left$(uni, 3) & " LOC  " & uni: nottun = False    'City
         Case 3: nottun = True: t_un = "SPDY " & Left$(uni, 3) & " LH  " & uni: nottun = False    'Linehaul
         Case 4: nottun = True: t_un = "SPDY " & Left$(uni, 3) & " " & uni: nottun = False    'unused
         Case 5: nottun = True: t_un = "SPDY " & Left$(uni, 3) & " SYR ST  " & uni: nottun = False    'Straight
         Case 6: nottun = True: t_un = "SPDY " & Left$(uni, 3) & " SYR " & uni: nottun = False    'unused
         Case 7 To 8: nottun = True: t_un = "SPDY " & Left$(uni, 3) & " " & uni: nottun = False    'unused
       End Select
       unit = uni
     Case Else
       nottun = True: t_un = BldUnit(uni): nottun = False
   End Select
 End If
 t_un.ForeColor = &H0
 'trailer=unit for straight & cube
 frun.Visible = False: fr1.Enabled = True
 If InStr(t_un, " ST ") > 0 Or InStr(t_un, " CUBE ") > 0 Then
   If Len(uni) = 8 Then trailer = uni Else trailer = Right$(uni, 4)
   If ch_htr Then ch_htr = False
   notttr = True: t_tr = trailer: notttr = False
   If trlrapp.Visible = True Then trlrapp.SetFocus
 Else
   If t_tr.Visible Then t_tr.SetFocus
 End If
End Sub
Private Sub lst_un_KeyDown(KeyCode%, Shf%)
 Select Case KeyCode
   Case vbKeyEscape: c_cancun_Click
   Case vbKeyReturn, vbKeyExecute: lst_un_DblClick
 End Select
End Sub
Private Sub lst_un_KeyPress(KeyAscii%)
 Select Case KeyAscii
   Case 48 To 57
   Case Else: KeyAscii = 0
 End Select
 If KeyAscii > 0 Then
   t_unfr = Chr$(KeyAscii): t_unfr.SetFocus
 End If
End Sub

Private Function BldUnit$(uni$)
 Dim dm As String, dm2 As String, unstr As String
 Dim n As Integer
 
 If IsNumeric(uni) = False Then GoTo ALPHUNITBU
 
 n = Val(uni)     'Val(Left$(lst_un.Text, 6))
 Select Case n
   Case Is < 200:
     Select Case n
       Case 6, 14, 16
       Case 77: dm2 = "LON"
       Case 38, 118, 139, 182: dm2 = "WND"
       Case 93, 149, 166: dm2 = "BRO"
       Case Else: dm2 = "TOR"
     End Select
   Case 500 To 600: dm2 = "MTL"
   Case 710, 720, 750, 764, 777, 7069: dm2 = "WND"
   Case 800 To 900: dm2 = "LON"
   Case Is > 1000
     Select Case n
       Case 1503, 6004, 6006, 6007, 7527, 7833, 8181, 9611: dm2 = "TOR"
       Case 6000: BldUnit = "CPR xTOR 6000": unit = "6000": Exit Function
       Case 6001: BldUnit = "CPR xMTL 6001": unit = "6001": Exit Function
       Case 5003, 5009 To 5013, 5015, 5017, 5019, 5024, 5026, 5027, 5029: dm2 = "MTL"
       Case 5031, 5033, 5035, 5041, 5043, 5047, 5051, 5054 To 5059: dm2 = "MTL"
       Case 5061, 5065, 5067, 5069, 5071, 5073, 5075, 5077: dm2 = "MTL"
       Case 8000 To 8999: dm2 = "LON"
       Case 1100 To 4099: dm2 = "BRO"
       Case Else: dm2 = "TOR"
     End Select
 End Select
 
'assigned MTL per Alain Contant eMail 15Nov16
'5-Ton
'5003
'5009
'5011
'5013
'5015
'5017
'5019
'5027
'5029
'5031
'5033
'5035
'5041
'5043
'5047
'5051
'5055
'5057
'5059
'5061
'5065
'5067
'5069
'5071
'5073
'5075
'5077
'Tractors:
'5010
'5012
'5024
'5026
'5054
'5056
'5058

 
 unit = Format$(n, "0000")
 If InStr(lst_un.Text, "Local") > 0 Then
   dm = "LOC "
 ElseIf Right$(lst_un.Text, 2) = "LH" Or InStr(lst_un.Text, " LH-") > 0 Then
   If InStr(lst_un.Text, "Montreal") > 0 Then
     dm = "MTLLH "
   ElseIf InStr(lst_un.Text, "Detroit") > 0 Then
     dm = "DETLH "
   ElseIf InStr(lst_un.Text, "Buffalo") > 0 Then
     dm = "BUFLH "
   ElseIf InStr(lst_un.Text, "Windsor") > 0 Then
     dm = "WINLH"
   Else
     dm = "XXXLH "
   End If
 Else
   dm = "  N/A "
 End If
 If InStr(lst_un.Text, "Straight") > 0 Then
   dm = dm & "  ST "
 ElseIf InStr(lst_un.Text, "Cube") > 0 Then
   dm = dm & "CUBE "
 Else
   If InStr(lst_un.Text, "Tand") > 0 Then
     dm = dm & "TNDM "
   ElseIf InStr(lst_un.Text, "Sing") > 0 Then
     dm = dm & "SNGL "
   End If
 End If
 BldUnit = "SPDY " & dm2 & " " & dm & unit
 
 Exit Function
 
ALPHUNITBU:
 unit = Trim$(uni)
 BldUnit = unit
 
End Function

Private Sub t_un_Change() 'user manually changes unit no.
 Dim dm As String
 Dim j As Integer
 Dim lu As String
 Dim s() As String
 
 If nottun Then Exit Sub

 dm = Trim$(t_un)

 '1. just enter any outside power unit as user-entered
 Select Case Trim$(t_uno)
   Case "", "SP", "SDT"
   Case Else
     unit = t_un: Exit Sub
 End Select
 
 '2. prevent legacy numeric unit entry from finding new style if legacy unit# does not exist
 Select Case Len(dm)
   Case Is < 3: Exit Sub
   Case 3
     If Left$(dm, 1) = "0" Then Exit Sub
 End Select
 
 '3. look for match in unit list - either by unit# segment or as old format #
 Select Case Left$(Trim$(t_un), 3)
   Case "TOR", "MTL", "LON", "PIC", "BRO", "WIN", "MIL"
     If Len(t_un) = 8 Then
       For j = 1 To lst_un.ListCount - 1
         s = Split(Trim$(lst_un.List(j)), " "): s(0) = Trim$(s(0))
         If t_un = s(0) Then
           unit = s(0): lst_un.ListIndex = j: t_un.ForeColor = &H0
           Select Case Mid(unit, 4, 1)
             Case 1 'straight, autoset trailer to same
               If Left(unit, 3) = Left$(Trim$(t_un), 3) Then
                 notttr = True: t_tr = unit:  trailer = unit: notttr = False
               End If
             Case 2 'city tractor
             Case 3 'highway tractor
             Case 4 'shunt truck
             Case 5 'unused
             Case 6 'Syracuse
             Case 7 'unused
             Case 8 'unused
           End Select
           Exit For
         End If
       Next j
     End If
   Case Else
     If Val(Trim$(t_un)) > 0 Then
       For j = 1 To lst_un.ListCount - 1 'sync hidden list with manually entered unit no.
         If Val(Trim$(t_un)) = Val(Left$(lst_un.List(j), 6)) Then 'match with old unit#
           unit = Format$(Val(t_un), "0000")
           lst_un.ListIndex = j: t_un.ForeColor = &H0 'black indicates unit no is good
           GoTo TUNC
         ElseIf IsNumeric(t_un) = True And Len(t_un) > 1 Then 'look for match with new unit#'s last 4 digit numeric code
           s = Split(Trim$(lst_un.List(j)), " "): s(0) = Trim$(s(0))
'           If s(0) = "TOR10421" Then
'             arg = arg
'           End If
           If t_un = Right$(s(0), Len(t_un)) And Val(t_un) = Val(Right(s(0), 4)) Then
             t_un = s(0): unit = s(0)
             lst_un.ListIndex = j: t_un.ForeColor = &H0: Exit For 'black indicates unit no is good
           End If
         End If
       Next j
     End If
 End Select
 unit = t_un
TUNC:
End Sub
Private Sub t_un_GotFocus()
 If g2profoc And g2pro.Visible Then
   If g2.Visible Then g2pro.SetFocus
   Exit Sub
 End If
 t_un.BackColor = &HFFFFFF
End Sub
Private Sub t_un_LostFocus() 'test unit no. manual entry on leaving
 Dim dm As String, uni As String
 Dim n As Long
 Dim s() As String
 
 t_un.BackColor = &HD6E6E6 'un-hilite field
 
 dm = ""
 'new format handling
 If Len(unit) = 8 Then
   Select Case Left(unit, 3)
      Case "TOR", "MTL", "LON", "PIC", "BRO", "WIN", "MIL"
        If IsNumeric(Mid(unit, 4)) = False Then
          If t_uno = "" Or t_uno = "SP" Or t_uno = "SDT" Then
            MsgBox "Invalid Format for New Unit Name. Please re-enter Unit (eg. MTL25005 = city tractor 5005 out of MTL terminal)", , "Trip Manifest"
            t_un.SetFocus: Exit Sub
          End If
        Else
          Select Case Val(Mid(unit, 4, 1))
            Case 1, 5 'straight
              notttr = True: t_tr = unit:  trailer = unit: notttr = False: Exit Sub
            Case 2 To 4, 6 To 8
            Case Else 'invalid
               MsgBox "Invalid Format of Unit Type Code in New Unit Name. Please re-enter Unit (eg. MTL25005, = city tractor 5005 out of MTL)", , "Trip Manifest"
               t_un.SetFocus: Exit Sub
          End Select
        End If
   End Select
 End If
 
 'find new format unit# with just the numeric part
 If IsNumeric(Trim$(t_uno)) = True Then
    uni = Trim$(t_uno)
    For j = 1 To lst_un.ListCount - 1
       dm = Replace(Trim$(lst_un.List(j)), "  ", " "): dm = Replace(Trim$(lst_un.List(j)), "  ", " ")
       s = Split(dm, " ")
       s(0) = Trim$(s(0)) 'unit#
       Select Case Left$(s(0), 3)
          Case "TOR", "MTL", "LON", "PIC", "BRO", "WIN", "MIL"
             If Right$(s(0), Len(uni)) = uni Then 'match with new unit# format
                nottun = True: t_un = s(0): unit = s(0): nottun = False
                'if straight truck, auto-enter the trailer
                If Mid$(unit, 4, 1) = "3" Then t_tr = unit
                Exit Sub
             End If
       End Select
    Next j
 End If

 t_un = Replace(Trim$(t_un), "  ", " ")
 If t_un = "" Then
   unit = "": Exit Sub
 End If
 s = Split(t_un, " "): uni = s(UBound(s))
 If Trim$(t_uno) <> "" And t_uno <> "SP" And t_uno <> "SDT" Then 'sdt*
   If Len(uni) > 8 Then MsgBox "Unit Entry Cannot Exceed 8 Characters When Outside-Owner is Selected!", , "Manifest"
   unit = Right$(t_un, 8): t_un = unit
   Exit Sub
 End If
 
 n = Val(uni)   'n = Val(Right$(t_un, 8)) 'get unit numeric value
 Select Case Left$(uni, 3)
   Case "TOR", "MTL", "LON", "PIC", "BRO", "WIN", "MIL"
     If Len(uni) <> 8 Then
       MsgBox "New Speedy Unit Names (eg. TOR20123) are Always 8-Chars", , "Manifest"
       unit = Right$(uni, 8): t_un = unit
       Exit Sub
     End If
     Select Case Mid$(uni, 4, 1)
       Case 1, 5 'if Speedy or Syracuse straight truck, get equivalent trailer
         If Len(uni) = 8 Then trailer = uni Else trailer = Right$(uni, 4)
         t_tro = "SP"
         If ch_htr Then ch_htr = False
         notttr = True: t_tr = trailer: notttr = False 'write trailer = unit
         t_tr.ForeColor = &H0
     End Select
     unit = uni
   Case Else
     If n = 0 Then
       If uni = "" Then
         unit = "": t_un = ""
       Else
         unit = uni
       End If
     Else
       For j = 1 To lst_un.ListCount - 1
         dm = Replace(Trim$(lst_un.List(j)), "  ", " ")
         s = Split(dm, " ")
         If n = Val(s(0)) Then  'match found
           lst_un.ListIndex = j: t_un.ForeColor = &H0
           nottun = True: t_un = BldUnit(uni): nottun = False 'write unit with all info
           If InStr(t_un, " ST ") > 0 Or InStr(t_un, " CUBE ") > 0 Then 'trailer = unit for straight & cube units
             trailer = unit: t_tro = "SP" 'sdt*
             If ch_htr Then ch_htr = False
             notttr = True: t_tr = unit: notttr = False 'write trailer = unit
             t_tr.ForeColor = &H0  '': trlrapp.SetFocus 03Feb11 bugs dropdown list!!
           Else 'not straight/cube
             If Len(t_tr) = 5 Then 'check current trailer entry (if any) - if exactly 5 chars then could be leftover straight/cube entry
               For n = 1 To lst_tr.ListCount - 1 'cycle thru trailer list
                 If t_tr = Trim$(Left$(lst_tr.Text, 10)) Then GoTo TUNLF1 'trailer is valid - do nothing
               Next n
               trailer = "": notttr = True: t_tr = "": notttr = False 'if here then trailer not valid - clear
TUNLF1:
             End If
             ''t_tr.SetFocus 03Feb11 bugs dropdown list!!
           End If
           GoTo TUNLF2 'unit no was valid
         End If
       Next j
       unit = t_un
     End If
 End Select
 
TUNLF2:
End Sub
Private Sub t_un_KeyDown(KeyCode%, Shf%) 'user key on unit no. field
 Select Case KeyCode
   Case vbKeyEscape: t_un = "": unit = ""
   Case vbKeyDown, vbKeyReturn, vbKeyExecute
     If InStr(t_un, " ST ") > 0 Then
       If trlrapp.Visible Then trlrapp.SetFocus
     Else
       If t_tr.Visible Then t_tr.SetFocus
     End If
   Case vbKeyPageUp, vbKeyPageDown: c_un_Click
 End Select
End Sub
Private Sub t_un_KeyPress(KeyAscii%) 'manual data entry keypress on unit no. field
 Select Case KeyAscii
   Case 8, 32, 48 To 57, 65 To 90 'bkspace, space, 0 - 9, A - Z
   Case 97 To 122: KeyAscii = KeyAscii - 32 'a-z convert to uppercase
   Case Else: KeyAscii = 0  'filter out all other chars
 End Select
 If Len(t_un) = 0 Then
   Select Case KeyAscii
     Case 35: t_un = "MTL": KeyAscii = 0: t_un.SelStart = 4
     Case 66: t_un = "BRO": KeyAscii = 0: t_un.SelStart = 4
     Case 73: t_un = "MIL": KeyAscii = 0: t_un.SelStart = 4 'I
     Case 76: t_un = "LON": KeyAscii = 0: t_un.SelStart = 4 'L
     'Case 77: t_un = "MTL": KeyAscii = 0: t_un.SelStart = 4 'cannot use M - MTL or MIL?
     Case 80: t_un = "PIC": KeyAscii = 0: t_un.SelStart = 4 'P
     Case 83: t_un = "MTL": KeyAscii = 0: t_un.SelStart = 4 'S for St-Laurent
     Case 84: t_un = "TOR": KeyAscii = 0: t_un.SelStart = 4 'T
     Case 87: t_un = "WIN": KeyAscii = 0: t_un.SelStart = 4 'W
     Case 91: t_un = "MTL": KeyAscii = 0: t_un.SelStart = 4
   End Select
 End If
End Sub

Private Sub t_htrfr_Change()
 Dim j As Integer
 If nothtrfr Then Exit Sub
 For j = 1 To lst_htr.ListCount - 1
   If t_htrfr = Trim$(Left$(lst_htr.List(j), 7)) Then
     lst_htr.ListIndex = j: Exit For
   End If
 Next j
End Sub
Private Sub t_htrfr_GotFocus()
 t_htrfr.SelStart = Len(t_htrfr)
End Sub
Private Sub t_htrfr_KeyDown(KeyCode%, Shf%)
 Select Case KeyCode
   Case vbKeyEscape: c_canchtr_Click
   Case vbKeyReturn, vbKeyExecute: lst_htr_DblClick
 End Select
End Sub
Private Sub t_htrfr_KeyPress(KeyAscii%)
 KeyAscii = NumFilt(KeyAscii, True, t_htrfr.SelStart)
End Sub

Private Sub lst_htr_Click()
 nothtrfr = True
 t_htrfr = Trim$(Left$(lst_htr.Text, 7))
 nothtrfr = False
End Sub
Private Sub lst_htr_DblClick()
 Dim dm As String, dm2 As String
 trailer = Trim$(Left$(lst_htr.Text, 7))
 notttr = True: t_tr = trailer: notttr = False
 t_tr.ForeColor = &H0
 frhtr.Visible = False: fr1.Enabled = True
 If trlrapp.Visible Then trlrapp.SetFocus
End Sub
Private Sub lst_htr_KeyDown(KeyCode%, Shf%)
 Select Case KeyCode
   Case vbKeyEscape: c_canchtr_Click
   Case vbKeyReturn, vbKeyExecute: lst_htr_DblClick
 End Select
End Sub
Private Sub lst_htr_KeyPress(KeyAscii%)
 Select Case KeyAscii
   Case 48 To 57
   Case Else: KeyAscii = 0
 End Select
 If KeyAscii > 0 Then
   t_htrfr = Chr$(KeyAscii)
   If t_htrfr.Visible Then t_htrfr.SetFocus
 End If
End Sub

Private Sub t_tr_Change() 'user manually changes trailer no.
 Dim j As Integer
 Dim n As Long
 Dim b As Boolean
 Dim s() As String
 
 If notttr Then Exit Sub
 
 s = Split(Trim$(t_tr), " ")
 If UBound(s) > 0 Then
   trailer = s(UBound(s)): Exit Sub
 Else
   trailer = t_tr
 End If
 
 Select Case Trim$(t_uno)
   Case "", "SP", "SDT"
     If Len(t_tr) >= 3 And IsNumeric(Right$(t_tr, 4)) = True And InStr(Right$(t_tr, 4), "E") = 0 Then
       dm = "": n = Right(t_tr, 4)
       Select Case n
         Case Is > 999: dm = CStr(n)
         Case Is > 99: dm = "0" & CStr(n)
       End Select
       If dm <> "" Then
          For j = 0 To lst_tr.ListCount - 1
            dm2 = Trim$(Mid$(lst_tr.List(j), 6, 8))
            Select Case Left$(dm2, 3)
              Case "TOR", "MTL", "LON", "PIC", "BRO", "WIN", "MIL"
                If Right$(dm2, 5) = "1" & dm Then
                  t_tr = dm2: t_un = dm2: Exit Sub
                End If
            End Select
          Next j
       End If
     End If
   Case Else
     trailer = t_tr: Exit Sub
 End Select
  
 If Trim$(t_tr) = "" Then Exit Sub
 
 If ch_htr Then 'US trailer
   If Len(t_tr) < 6 Then
     t_tr.ForeColor = &H80
   Else
     For j = 0 To lst_htr.ListCount - 1
       If t_tr = Trim$(Left$(lst_htr.List(j), 7)) Then
         trailer = t_tr: t_tr = "US " & t_tr: t_tr.ForeColor = &H0: GoTo TTRB1
       End If
     Next j
     trailer = "": t_tr.ForeColor = &H80
TTRB1:
   End If
   Exit Sub
 End If
 
 
 
 If InStr(t_tr, "E") > 0 Then
   n = 0: b = False
 Else
   n = Val(Trim$(t_tr)): b = IsNumeric(Trim$(t_tr))
 End If
 For j = 1 To lst_tr.ListCount - 1
   If Len(t_tr) = 8 Then
     Select Case Left$(t_tr, 3)
       Case "TOR", "MTL", "LON", "PIC", "BRO", "WIN", "MIL"
         If Mid(t_tr, 4, 1) = 1 And IsNumeric(Mid(t_tr, 4)) = True Then
           If Right$(t_un, 8) <> t_tr Then t_un = t_tr 'correct unit for straight truck selection
         End If
     End Select
   End If
  
   If Not b Then
     If t_tr = Trim$(Mid$(lst_tr.List(j), 3, 11)) Then
       lst_tr.ListIndex = j
       trailer = Trim$(Mid$(lst_tr.Text, 3, 11))
       t_tr.ForeColor = &H0: GoTo TTRB2
     End If
   Else
     If n = Val(Trim$(Mid$(lst_tr.List(j), 3, 11))) Then
       lst_tr.ListIndex = j
       trailer = Trim$(Mid$(lst_tr.Text, 3, 11))
       t_tr.ForeColor = &H0: GoTo TTRB2
     End If
   End If
 Next j

TTRB2:

End Sub
Private Sub t_tr_GotFocus()
 If g2profoc And g2pro.Visible Then
   If g2.Visible Then g2pro.SetFocus
   Exit Sub
 End If
 t_tr.BackColor = &HFFFFFF
End Sub
Private Sub t_tr_LostFocus() 'test trailer no.
 Dim dm As String, dm2 As String, dm3 As String
 Dim sp() As String
 Dim j As Integer, i As Integer
 Dim n As Long
 Dim b As Boolean
 
 g2profoc = False
 
 t_tr.BackColor = &HD6E6E6
 t_tr.ForeColor = &H0
 
 If Trim$(t_tr) = "" Then
   trailer = "": Exit Sub
 End If
 
 'look for owner-SCAC prefixed trailer# - process as-is
 If IsNumeric(Left(t_tr, 4)) = False Then
   dm2 = Left$(t_tr, 4)
   For j = 0 To lst_tro.ListCount - 1
     sp = Split(lst_tro.List(j), Chr(9))
     If dm2 = Trim$(sp(0)) Then
       lst_tro.ListIndex = j: lst_tro.Selected(j) = True: t_tro = dm2: trailer = Mid$(t_tr, 5): GoTo STTTR
     End If
   Next j
 End If
 
 If Trim$(t_tro) <> "" And t_tro <> "SP" Then
   t_tr = Trim$(t_tr)
   If Len(t_tr) > 8 Then MsgBox "Trailer Entry Cannot Exceed 8 Characters When Outside-Owner is Selected!", , "Manifest"
   trailer = Right$(t_tr, 8): t_tr = trailer
   Exit Sub
 End If
 
 'special ODFL trailers 'loaned?' for city use
 Select Case Left$(t_tr, 6)
   Case "531154", "531303", "531309", "531326", "531368"
     t_tro = "ODF": trailer = Left$(t_tr, 6)
     Exit Sub
 End Select
  
STTTR:
  
 'check if unit is straight or cubed
 If InStr(t_un, " ST ") > 0 Or InStr(t_un, "CUBE") > 0 Then
   t_tr.ForeColor = &H0: trailer = unit: t_tr = trailer 'force trailer=unit
   If Not chksvcno Then ChkService '#@#
   Exit Sub
 End If
 
 Select Case t_tro
   Case "", "SP"
     t_tr = Trim$(t_tr)
     If Len(t_tr) < 5 Then
       If IsNumeric(t_tr) = True Then t_tr = Format$(Val(t_tr), "0000")
     End If
 End Select
 
 If ch_htr Then 'US trailer
   If Left$(t_tr, 2) <> "US" Then
     n = Val(Trim$(Right$(t_tr, 6)))
     If n > 99999 And n < 1000000 Then 'trailer is in US 6-digit trailer no. range
       For j = 0 To lst_htr.ListCount - 1
         If n = Val(Trim$(Left$(lst_htr.List(j), 7))) Then
           notttr = True: t_tr = "US " & t_tr: notttr = False
           t_tr.ForeColor = &H0: GoTo TTRLF1
         End If
       Next j
       trailer = t_tr: notttr = True: t_tr = "US " & t_tr: notttr = False: t_tr.ForeColor = &H0
       Exit Sub
     Else 'not in current US trailer range
       If Trim$(t_tr) = "" Then 'empty
         trailer = ""
       Else
         ch_htr = False: trailer = t_tr: GoTo STTRLF 'unset US trailer - got Speedy trailer check
       End If
     End If
   Else
     If Left$(t_tr, 3) = "US " Then
       trailer = Trim$(Mid$(t_tr, 4))
     Else
       trailer = t_tr
     End If
   End If
TTRLF1:
   Exit Sub
 End If

STTRLF:

   If InStr(t_tr, "E") > 0 Then
     b = False
   Else
     b = IsNumeric(Trim$(t_tr))
   End If
   
   For j = 1 To lst_tr.ListCount - 1
     If Not b Then
       If t_tr = Trim$(Mid$(lst_tr.List(j), 3, 11)) Then
         lst_tr.ListIndex = j: t_tr.ForeColor = &H0: lst_tr_DblClick: Exit Sub
         t_tr.ForeColor = &H0: GoTo TTRLF2
       End If
     Else
       n = Val(Trim$(t_tr))
       If n = Val(Trim$(Mid$(lst_tr.List(j), 3, 11))) Then
         lst_tr.ListIndex = j: t_tr.ForeColor = &H0: lst_tr_DblClick: Exit Sub
         t_tr.ForeColor = &H0: GoTo TTRLF2
       End If
     End If
   Next j
TTRLF2:
 If Not chksvcno Then ChkService

End Sub
Private Sub ChkService()
 Dim dm As String
 Dim j As Integer
 
 Select Case t_tro
   Case "", "SP"
     dm = "select svccode from trlrsvc where trailer='" & trailer & "' "
     If Len(trailer) = 4 Then
       dm = dm & "or trailer='ST" & trailer & "' "
     End If
     dm = dm & "and owner='SZTG'"
     If rsc.State <> 0 Then rsc.Close
     rsc.Open dm
     If Not rsc.EOF Then
            Select Case rsc(0)
              Case "PM"
                dm = "** Trailer is Due for ** PREVENTIVE MAINTENANCE ** Before the End of This Month **" & _
                      vbCrLf & vbCrLf & "                                                  Please Dispatch Wisely"
                MsgBox dm, , "Trailer * PREVENTIVE MAINTENANCE * Advisory"
              Case "PO"
                dm = "** Trailer is ** OVERDUE! for PREVENTIVE MAINTENANCE! **" & _
                      vbCrLf & vbCrLf & "                                                  Please Dispatch Wisely"
                MsgBox dm, , "Trailer * OVERDUE! PREVENTIVE MAINTENANCE * Advisory"
              Case "AN"
                dm = "** Trailer is Due for ** ANNUAL MAINTENANCE ** Before the End of This Month **" & _
                      vbCrLf & vbCrLf & "                                               Please Dispatch Wisely"
                MsgBox dm, , "Trailer * ANNUAL MAINTENANCE * Advisory"
              Case "RR"
                dm = "  ** Trailer is Due for ** REPAIR ASAP **" & _
                      vbCrLf & vbCrLf & "** Please Keep to Local Trip Only If Used **"
                MsgBox dm, , "Trailer * REPAIR REQ'D * Advisory"
              Case "OS"
                dm = "     ** Trailer is ** OUT-OF-SERVICE ** DO NOT DISPATCH **" & _
                      vbCrLf & vbCrLf & "                              Please Use Alternate Equipment"
                MsgBox dm, , "Trailer * OUT-OF-SERVICE * DO NOT DISPATCH * Advisory"
                t_tr = ""
           End Select
     End If
     rsc.Close
 End Select
End Sub

Private Sub t_tr_KeyDown(KeyCode%, Shf%) 'user key on trailer no. field
 Select Case KeyCode
   Case vbKeyEscape: t_tr = "": trailer = ""
   Case vbKeyUp
     If t_un.Visible Then t_un.SetFocus
   Case vbKeyDown, vbKeyReturn, vbKeyExecute
     If trlrapp.Visible Then trlrapp.SetFocus
   Case vbKeyPageUp, vbKeyPageDown: c_tr_Click
 End Select
End Sub
Private Sub t_tr_KeyPress(KeyAscii%) 'manual data entry keypress on trailer no. field
 If ch_htr Then t_tr = Replace(t_tr, "US ", "")
 Select Case KeyAscii
   Case 8, 32, 48 To 57, 65 To 90 'bkspace, 0 - 9, A - Z
   Case 97 To 122: KeyAscii = KeyAscii - 32 'a-z convert to uppercase
   Case Else: KeyAscii = 0  'filter out all other chars
 End Select
End Sub


Private Sub t_trfr_Change() 'change to list-find field above speedy trailers list
 Dim j As Integer
 If nottrfr Then Exit Sub
 If t_trfr = "" Then
   lst_tr.ListIndex = 0
 Else
   For j = 1 To lst_tr.ListCount - 1
     If t_trfr = Trim$(Mid$(lst_tr.Text, 3, 11)) Then
       lst_tr.ListIndex = j: Exit For
     End If
   Next j
 End If
End Sub
Private Sub t_trfr_GotFocus()
 t_trfr.SelStart = Len(t_trfr)
End Sub
Private Sub t_trfr_KeyDown(KeyCode%, Shf%)
 Select Case KeyCode
   Case vbKeyEscape: c_canctr_Click
   Case vbKeyReturn, vbKeyExecute
     If lst_tr.ListIndex = 0 Then
       If Trim$(t_trfr) = "" Then lst_tr_DblClick
     Else
       If Trim$(Mid$(lst_tr.Text, 3, 11)) = Trim$(t_trfr) Then lst_tr_DblClick
     End If
 End Select
End Sub
Private Sub t_trfr_KeyPress(KeyAscii%)
 Select Case KeyAscii
   Case 8, 48 To 57, 65 To 90 'bkspace, 0 - 9, A - Z
   Case 97 To 122: KeyAscii = KeyAscii - 32 'a-z convert to uppercase
   Case Else: KeyAscii = 0  'filter out all other chars
 End Select
End Sub

Private Sub typstor_Click() 'mark/un-mark trailer as Storage
 Dim dm As String, dm2 As String, dom As String, tr As String
 Dim nt As Boolean
  
 dm2 = " " & Trim$(Mid$(lst_tr.Text, 3)): tr = Trim$(Left$(dm2, 9))
 If rs.State <> 0 Then rs.Close
 rs.Open "select storage,domestic from trlrtypehistory where trailer='" & tr & "' and owner='SZTG' order by recno desc limit 1"
 If rs.EOF Then
   nt = True: dom = ""
 Else
   If Val(rs(0)) = 1 Then nt = False Else nt = True
   dom = CStr(rs(1))
 End If
 rs.Close
 If nt Then
   dm = "Confirm - * SET * Owner/Trailer: " & cscac & dm2 & "  as STORAGE TRAILER ?"
   If MsgBox(dm, vbDefaultButton2 + vbQuestion + vbYesNo, "* SET * Trailer " & dm2 & " for STORAGE USE ONLY") = vbNo Then Exit Sub
   dm2 = "1"
 Else
   dm = "Confirm - * UN-SET * Owner/Trailer: " & cscac & dm2 & "  as STORAGE TRAILER ?"
   If MsgBox(dm, vbDefaultButton1 + vbQuestion + vbYesNo, "* UN-SET * Trailer " & dm2 & " for STORAGE USE ONLY") = vbNo Then Exit Sub
   dm2 = ""
 End If
 dm = "insert ignore into trlrtypehistory set trailer='" & tr & "',owner='" & cscac & "',storage='" & dm2
 dm = dm & "',domestic='" & dom & "',initstorage='" & init & "',setdtstorage='" & Format$(Now, "YYYY-MM-DD Hh:Nn:Ss") & "'"
 cmd.CommandText = dm: cmd.Execute
 forcetrailerlistrefresh = True: tmr_Timer: forcetrailerlistrefresh = False
End Sub

Private Sub typdom_Click() 'mark/un-mark trailer as Domestic
 Dim dm As String, dm2 As String, stor As String, tr As String
 Dim nt As Boolean
  
 dm2 = " " & Trim$(Mid$(lst_tr.Text, 3)): tr = Trim$(Left$(dm2, 9))
 If rs.State <> 0 Then rs.Close
 rs.Open "select storage,domestic from trlrtypehistory where trailer='" & tr & "' and owner='SZTG' order by recno desc limit 1"
 If rs.EOF Then
   nt = True: stor = ""
 Else
   If Val(rs(1)) = 1 Then nt = False Else nt = True
   stor = CStr(rs(0))
 End If
 rs.Close
 If nt Then
   dm = "Confirm - * SET * Owner/Trailer: " & cscac & dm2 & "  as DOMESTIC ONLY TRAILER ?"
   If MsgBox(dm, vbDefaultButton2 + vbQuestion + vbYesNo, "* SET * Trailer " & dm2 & " for DOMESTIC USE ONLY") = vbNo Then Exit Sub
   dm2 = "1"
 Else
   dm = "Confirm - * UN-SET * Owner/Trailer: " & cscac & dm2 & "  as DOMESTIC ONLY TRAILER ?"
   If MsgBox(dm, vbDefaultButton1 + vbQuestion + vbYesNo, "* UN-SET * Trailer " & dm2 & " for DOMESTIC USE ONLY") = vbNo Then Exit Sub
   dm2 = ""
 End If
 dm = "insert ignore into trlrtypehistory set trailer='" & tr & "',owner='" & cscac & "',storage='" & stor
 dm = dm & "',domestic='" & dm2 & "',initdomestic='" & init & "',setdtdomestic='" & Format$(Now, "YYYY-MM-DD Hh:Nn:Ss") & "'"
 cmd.CommandText = dm: cmd.Execute
 forcetrailerlistrefresh = True: tmr_Timer: forcetrailerlistrefresh = False
End Sub

Private Sub trpm_Click()
 Dim dm As String, dm2 As String, dm3 As String, str As String, ow As String
 Dim k As Integer
 
 dm2 = Trim$(Mid$(clstc.Text, 3, 11))
 dm = "Confirmation: MARK " & cownc & " Trailer " & dm2 & " for PREVENTIVE MAINTENANCE ?"
 If MsgBox(dm, vbDefaultButton2 + vbQuestion + vbYesNo, "Service Advisory Confirmation") = vbNo Then Exit Sub
 str = "PM" & Mid$(clstc.Text, 3): k = clstc.ListIndex 'buff list items
 'trlrsvc - records created with service code, then deleted when returned to regular service
 ' recno   | int(8) unsigned    |
 ' trailer | varchar(16) binary |
 ' owner   | char(4) binary     |
 ' svccode | char(2) binary     |
 ' descrip | varchar(64) binary |
 ' init    | char(3)            |
 ' dt      | datetime           |
 If rs.State <> 0 Then rs.Close
 rs.Open "select recno from trlrsvc where trailer='" & Trim$(Mid$(clstc.Text, 3, 11)) & "' and owner='" & cscac & "' order by recno desc limit 1"
 If rs.EOF Then
   'create service advisory record
   dm = "insert ignore into trlrsvc set trailer='" & Trim$(Mid$(clstc.Text, 3, 11)) & "',owner='" & cscac & _
        "',svccode='PM',init='" & init & "', dt='" & Format$(Now, "YYYY-MM-DD Hh:Nn") & ":00" & "', detail='Set-PM'"
 Else
   'update existing advisory record
   dm = "update trlrsvc set svccode='PM',init='" & init & "', dt='" & Format$(Now, "YYYY-MM-DD Hh:Nn") & ":00' where recno='" & rs(0) & "'"
 End If
 rs.Close
 On Error Resume Next
 cmd.CommandText = dm: Err = 0: cmd.Execute
 If Err <> 0 Then
   MsgBox Err.Description, , "Service Advisory"
'trlrstat
' recno   | bigint(12) unsigned |
' owner   | varchar(4) binary   |
' trailer | varchar(16) binary  |
' created | datetime            |
' source  | char(1) binary      |
' init    | char(3) binary      |
' status  | varchar(64) binary  |
' code    | char(3) binary      |
 Else
   'add trailer status record
   dm = "insert ignore into trlrstat set trailer='" & Trim$(Mid$(clstc.Text, 3, 11)) & "',owner='" & cscac & "',init='" & _
         init & "', created='" & Format$(Now, "YYYY-MM-DD Hh:Nn") & ":00" & "', source='T',status='Set-PM',code='SPM'"
   cmd.CommandText = dm: Err = 0: cmd.Execute
   If Err <> 0 Then
     MsgBox Err.Description, , "Service Advisory": Exit Sub
   End If
   '13Nov20------
   dm = cscac & " Trailer " & dm2 & " PM Preventative Maintenance - SET (" & init & ") " & Format$(Now, "MMM-DD Hh:Nn")
   SendMail dm, vbCrLf & dm, "svctrpm"
   '-------------
   clstc.RemoveItem k: clstc.AddItem str, k
   forcetrailerlistrefresh = True: tmr_Timer: forcetrailerlistrefresh = False
 End If
End Sub

Private Sub trpo_Click()
 Dim dm As String, dm2 As String, dm3 As String, str As String, ow As String
 Dim k As Integer
 
 dm2 = Trim$(Mid$(clstc.Text, 3, 11))
 dm = "Confirmation: MARK " & cownc & " Trailer " & dm2 & " for OVERDUE! PREVENTIVE MAINTENANCE ?"
 If MsgBox(dm, vbDefaultButton2 + vbQuestion + vbYesNo, "Service Advisory Confirmation") = vbNo Then Exit Sub
 str = "PO" & Mid$(clstc.Text, 3): k = clstc.ListIndex 'buff list items
 If rs.State <> 0 Then rs.Close
 rs.Open "select recno from trlrsvc where trailer='" & Trim$(Mid$(clstc.Text, 3, 11)) & "' and owner='" & cscac & "'"
 If rs.EOF Then
   dm = "insert ignore into trlrsvc set trailer='" & Trim$(Mid$(clstc.Text, 3, 11)) & "',owner='" & cscac & _
        "',svccode='PO',init='" & init & "', dt='" & Format$(Now, "YYYY-MM-DD Hh:Nn") & ":00" & "', detail='Set-PO'"
 Else
   dm = "update trlrsvc set svccode='PO',init='" & init & "', dt='" & Format$(Now, "YYYY-MM-DD Hh:Nn") & ":00' where recno='" & rs(0) & "'"
 End If
 rs.Close
 On Error Resume Next
 cmd.CommandText = dm: Err = 0: cmd.Execute
 If Err <> 0 Then
   MsgBox Err.Description, , "Service Advisory"
 Else
   dm = "insert ignore into trlrstat set trailer='" & Trim$(Mid$(clstc.Text, 3, 11)) & "',owner='" & cscac & "',init='" & _
         init & "', created='" & Format$(Now, "YYYY-MM-DD Hh:Nn") & ":00" & "', source='T',status='Set-PO',code='SPO'"
   cmd.CommandText = dm: Err = 0: cmd.Execute
   If Err <> 0 Then
     MsgBox Err.Description, , "Service Advisory": Exit Sub
   End If
   '13Nov20------
   dm = cscac & " Trailer " & dm2 & " PO Preventative Maintenance Overdue - SET (" & init & ") " & Format$(Now, "MMM-DD Hh:Nn")
   SendMail dm, vbCrLf & dm, "svctrpo"
   '-------------
   clstc.RemoveItem k: clstc.AddItem str, k
   forcetrailerlistrefresh = True: tmr_Timer: forcetrailerlistrefresh = False
 End If
End Sub


Private Sub trrr_Click()
 Dim dm As String, dm2 As String, dm3 As String, str As String, ow As String
 Dim k As Integer
 
 dm2 = Trim$(Mid$(clstc.Text, 3, 11))
 dm = "Confirmation: MARK " & cownc & " Trailer " & dm2 & " for REPAIR REQ'D ?"
 If MsgBox(dm, vbDefaultButton2 + vbQuestion + vbYesNo, "Service Advisory Confirmation") = vbNo Then Exit Sub
 str = "RR" & Mid$(clstc.Text, 3): k = clstc.ListIndex
 If rs.State <> 0 Then rs.Close
 rs.Open "select recno from trlrsvc where trailer='" & Trim$(Mid$(clstc.Text, 3, 11)) & "' and owner='" & cscac & "'"
 If rs.EOF Then
   dm = "insert ignore into trlrsvc set trailer='" & Trim$(Mid$(clstc.Text, 3, 11)) & "',owner='" & cscac & _
        "',svccode='RR',init='" & init & "', dt='" & Format$(Now, "YYYY-MM-DD Hh:Nn") & ":00" & "', detail='Set-RR'"
 Else
   dm = "update trlrsvc set svccode='RR',init='" & init & "', dt='" & Format$(Now, "YYYY-MM-DD Hh:Nn") & ":00' where recno='" & rs(0) & "'"
 End If
 rs.Close
 On Error Resume Next
 cmd.CommandText = dm: Err = 0: cmd.Execute
 If Err <> 0 Then
   MsgBox Err.Description, , "Service Advisory"
 Else
   dm = "insert ignore into trlrstat set trailer='" & Trim$(Mid$(clstc.Text, 3, 11)) & "',owner='" & cscac & "',init='" & _
         init & "', created='" & Format$(Now, "YYYY-MM-DD Hh:Nn") & ":00" & "', source='T',status='Set-RR',code='SRR'"
   cmd.CommandText = dm: Err = 0: cmd.Execute
   If Err <> 0 Then
     MsgBox Err.Description, , "Service Advisory": Exit Sub
   End If
   '16Nov20------
   dm = cscac & " Trailer " & dm2 & " RR Repair Required - SET (" & init & ") " & Format$(Now, "MMM-DD Hh:Nn")
   SendMail dm, vbCrLf & dm, "svctrrr"
   '-------------
   clstc.RemoveItem k: clstc.AddItem str, k
   forcetrailerlistrefresh = True: tmr_Timer: forcetrailerlistrefresh = False
 End If
End Sub

Private Sub tran_Click()
 Dim dm As String, dm2 As String, dm3 As String, str As String, ow As String
 Dim k As Integer
 
 dm2 = Trim$(Mid$(clstc.Text, 3, 11))
 dm = "Confirmation: MARK " & cownc & " Trailer " & dm2 & " for ANNUAL MAINTENANCE ?"
 If MsgBox(dm, vbDefaultButton2 + vbQuestion + vbYesNo, "Service Advisory Confirmation") = vbNo Then Exit Sub
 str = "AN" & Mid$(clstc.Text, 3): k = clstc.ListIndex
 If rs.State <> 0 Then rs.Close
 rs.Open "select recno from trlrsvc where trailer='" & Trim$(Mid$(clstc.Text, 3, 11)) & "' and owner='" & cscac & "'"
 If rs.EOF Then
   dm = "insert ignore into trlrsvc set trailer='" & Trim$(Mid$(clstc.Text, 3, 11)) & "',owner='" & cscac & _
        "',svccode='AN',init='" & init & "', dt='" & Format$(Now, "YYYY-MM-DD Hh:Nn") & ":00" & "', detail='Set-AN'"
 Else
   dm = "update trlrsvc set svccode='AN',init='" & init & "', dt='" & Format$(Now, "YYYY-MM-DD Hh:Nn") & ":00' where recno='" & rs(0) & "'"
 End If
 rs.Close
 On Error Resume Next
 cmd.CommandText = dm: Err = 0: cmd.Execute
 If Err <> 0 Then
   MsgBox Err.Description, , "Service Advisory"
 Else
   dm = "insert ignore into trlrstat set trailer='" & Trim$(Mid$(clstc.Text, 3, 11)) & "',owner='" & cscac & "',init='" & _
         init & "', created='" & Format$(Now, "YYYY-MM-DD Hh:Nn") & ":00" & "', source='T',status='Set-AN',code='SAN'"
   cmd.CommandText = dm: Err = 0: cmd.Execute
   If Err <> 0 Then
     MsgBox Err.Description, , "Service Advisory": Exit Sub
   End If
   clstc.RemoveItem k: clstc.AddItem str, k
   '12Nov20------
   dm = cscac & " Trailer " & dm2 & " AN Annual Maintenance - SET (" & init & ") " & Format$(Now, "MMM-DD Hh:Nn")
   SendMail dm, vbCrLf & dm, "svctran"
   '-------------
    'Arlene Short
    'Jennifer Cole
    'Leana Cadeau
 End If
End Sub


Private Sub trrem_Click()
 Dim dm As String, dm2 As String, dm3 As String, str As String, oldcode As String, oldrec As String
 Dim sq As String, detail As String
 Dim j As Integer
 
 dm2 = Trim$(Mid$(clstc.Text, 3, 11))
 dm = "Confirmation: REMOVE Service Advisory from " & cownc & " Trailer " & dm2 & " ?"
 If MsgBox(dm, vbDefaultButton2 + vbQuestion + vbYesNo, "Remove Service Advisory Confirmation") = vbNo Then Exit Sub
 
 '12Nov20 -------------------------
 sq = ""
 If rs.State <> 0 Then rs.Close
 dm = "select cast(dt as char) as ""dtset"", init, svccode, recno, detail from trlrsvc where trailer='" & dm2 & _
      "' and owner='" & cscac & "' order by dt desc limit 1"
 rs.Open dm
 If Not rs.EOF Then
   detail = Trim$(rs!detail)
   If detail = "" And rs!svccode = "OS" Then
     If rsb.State <> 0 Then rsb.Close
     dm = "select status from trlrstat where trailer='" & dm2 & "' and owner='" & cscac & "' and code='SOS' " & _
          "and left(cast(created as char), 16)='" & Left$(rs!dtset, 16) & "'"
     rsb.Open dm
     If Not rsb.EOF Then detail = rsb!Status
     rsb.Close
   End If
   sq = "insert ignore into trlrsvcrmv set trailer='" & dm2 & "', owner='" & cscac & "', dtset='" & rs!dtset & "', setinit='" & rs!init & _
        "', svccode='" & rs!svccode & "', dtremoved='" & Format$(Now, "YYYY-MM-DD Hh:Nn:Ss") & "', removedinit='" & init & _
        "', setrecno='" & rs!recno & "', detail='" & mySav(detail) & "'"
   rs.Close
   On Error Resume Next
   cmd.CommandText = sq: Err = 0: cmd.Execute
   If Err <> 0 Then MsgBox Err.Description
 End If
 If rs.State <> 0 Then rs.Close
 '-------------------------------
 
 str = "  " & Mid$(clstc.Text, 3): k = clstc.ListIndex
 dm = "delete from trlrsvc where trailer='" & dm2 & "' and owner='" & cscac & "'"
 On Error Resume Next
 cmd.CommandText = dm: Err = 0: cmd.Execute
 If Err <> 0 Then
   MsgBox Err.Description, , "Service Advisory"
 Else
   If Left$(clstc.Text, 2) = "OS" Then 'was OS, update status, send mail
     dm = "insert ignore into trlrstat set trailer='" & dm2 & "',owner='" & cscac & "',init='" & _
           init & "', created='" & Format$(Now, "YYYY-MM-DD Hh:Nn") & ":00" & "', source='T',status='" & mySav(gs1) & "',code='SVC'"
     cmd.CommandText = dm: Err = 0: cmd.Execute
     If Err <> 0 Then
       MsgBox Err.Description, , "Service Advisory"
     Else
       ''dm = cownc & " Trailer " & dm2 & " Returned to Service " & init & " " & Format$(Now, "DD-MMMYY Hh:Nnam/pm")
       dm = cscac & " Trailer " & dm2 & " Returned to Service " & init & " " & Format$(Now, "DD-MMMYY Hh:Nnam/pm")
       SendMail dm, vbCrLf & dm, "svctros"
     End If
     clstc.RemoveItem k: clstc.AddItem str, k
   Else
     'DBUG Exit Sub
     dm = ""
     Select Case Left$(clstc.Text, 2)
       Case "AN"
         dm = cscac & " Trailer " & dm2 & " AN Annual Maintenance - CLEARED (" & init & ") " & Format$(Now, "MMM-DD Hh:Nn")
         SendMail dm, vbCrLf & dm, "svctran"
       Case "PM"
         dm = cscac & " Trailer " & dm2 & " PM Preventative Maintenance - CLEARED (" & init & ") " & Format$(Now, "MMM-DD Hh:Nn")
         SendMail dm, vbCrLf & dm, "svctrpm"
       Case "PO"
         dm = cscac & " Trailer " & dm2 & " PO Preventative Maintenance Overdue - CLEARED (" & init & ") " & Format$(Now, "MMM-DD Hh:Nn")
         SendMail dm, vbCrLf & dm, "svctrpo"
       Case "RR"
         dm = cscac & " Trailer " & dm2 & " RR Repaired Required - CLEARED (" & init & ") " & Format$(Now, "MMM-DD Hh:Nn")
         SendMail dm, vbCrLf & dm, "svctrrr"
     End Select
   End If
   forcetrailerlistrefresh = True: tmr_Timer: forcetrailerlistrefresh = False
 End If
End Sub

Private Sub tros_Click()
 Dim dm As String, dm2 As String, str As String
 
 gs1 = "": gs2 = Trim$(Mid$(clstc.Text, 3, 11)): gs3 = "SP"
 svcfrm.Show 1
 If gs1 = "!X@" Then Exit Sub 'cancel
 If Trim$(gs1) = "" Then gs1 = "Set-OS"
 str = "OS" & Mid$(clstc.Text, 3): k = clstc.ListIndex
 
 If rs.State <> 0 Then rs.Close
 rs.Open "select recno from trlrsvc where trailer='" & Trim$(Mid$(clstc.Text, 3, 11)) & "' and owner='" & cscac & "'"
 If rs.EOF Then
   dm = "insert ignore into trlrsvc set trailer='" & Trim$(Mid$(clstc.Text, 3, 11)) & "',owner='" & cscac & _
        "',svccode='OS', init='" & init & "', dt='" & Format$(Now, "YYYY-MM-DD Hh:Nn") & ":00" & "', detail='" & mySav(gs1) & "'"
 Else
   dm = "update trlrsvc set svccode='OS',init='" & init & "', dt='" & Format$(Now, "YYYY-MM-DD Hh:Nn") & ":00' where recno='" & rs(0) & "'"
 End If
 rs.Close
 On Error Resume Next
 cmd.CommandText = dm: Err = 0: cmd.Execute
 If Err <> 0 Then
   MsgBox Err.Description, , "Service Advisory"
 Else
   dm = "insert ignore into trlrstat set trailer='" & Trim$(Mid$(clstc.Text, 3, 11)) & "',owner='" & cscac & "',init='" & _
         init & "', created='" & Format$(Now, "YYYY-MM-DD Hh:Nn") & ":00" & "', source='T',status='" & mySav(gs1) & "',code='SOS'"
   cmd.CommandText = dm: Err = 0: cmd.Execute
   If Err <> 0 Then
     MsgBox Err.Description, , "Service Advisory"
   Else
     'send mail on set out-of-service
     ''dm = cownc & " Trailer " & gs2 & " Set to Out-of-Service " & init & " " & Format$(Now, "DD-MMMYY Hh:Nnam/pm")
     dm = cscac & " Trailer " & gs2 & " Set to Out-of-Service " & init & " " & Format$(Now, "DD-MMMYY Hh:Nnam/pm")
     dm2 = vbCrLf & dm
     If gs1 <> "" Then dm2 = dm2 & vbCrLf & vbCrLf & "Detail: " & gs1
     SendMail dm, dm2, "svctros"
     clstc.RemoveItem k: clstc.AddItem str, k
   End If
   forcetrailerlistrefresh = True: tmr_Timer: forcetrailerlistrefresh = False
 End If
End Sub



Private Sub lst_tr_MouseDown(Button%, Shf%, x!, y!)
 Dim dm As String, dm2 As String, tr As String
 Dim j As Integer
 
 If Button <> 2 Then Exit Sub
 If lst_tr.Text = "" Or Left$(lst_tr.Text, 7) = "   - UN" Then Exit Sub
 
 typstor.Caption = "Mark Trailer: STORAGE-ONLY": typdom.Caption = "Mark Trailer: DOMESTIC-ONLY"
 dm2 = " " & Trim$(Mid$(lst_tr.Text, 3)): tr = Trim$(Left$(dm2, 9))
 If rs.State <> 0 Then rs.Close
 rs.Open "select storage,domestic from trlrtypehistory where trailer='" & tr & "' and owner='SZTG' order by recno desc limit 1"
 If Not rs.EOF Then
   If Val(rs(0)) = 1 Then typstor.Caption = "*UN-MARK* Trailer: Storage-Only"
   If Val(rs(1)) = 1 Then typdom.Caption = "*UN-MARK* Trailer: Domestic-Only"
 End If
 rs.Close
 
 Select Case Val(utsvc)
   Case 1 'set advisory menu
     trti.Caption = "Speedy Trailer: " & Trim$(Mid$(lst_tr.Text, 3, 11)) & " Service Advisory" '/Status History"
     s16.Visible = True: s17.Visible = True: s18.Visible = True: s28.Visible = True: s29.Visible = True  '   : s19.Visible = True: s20.Visible = True
     trpm.Visible = True: tran.Visible = True: trrr.Visible = True: tros.Visible = True: trrem.Visible = True
     trpm.Enabled = True: tran.Enabled = True: trrr.Enabled = True: tros.Enabled = True: trrem.Enabled = True
     Select Case Trim$(Left$(lst_tr.Text, 2))
       Case "": trrem.Enabled = False
       Case "PM": trpm.Enabled = False
       Case "AN": tran.Enabled = False
       Case "RR": trrr.Enabled = False
       Case "OS": tros.Enabled = False
     End Select
   Case Else 'show status history select
     Exit Sub '<---
 End Select
 Set clstc = lst_tr: cownc = "Speedy": cscac = "SZTG"
 PopupMenu trpop
End Sub
Private Sub lst_tr_Click() 'change hilited item in speedy trailer list
 nottrfr = True
  If lst_tr.ListIndex = 0 Then t_trfr = "" Else t_trfr = Trim$(Mid$(lst_tr.Text, 3, 11))
  nottrfr = False
End Sub
Private Sub lst_tr_DblClick() 'select item from speedy trailer list
 Dim dm As String, dm2 As String
 Dim j As Integer
 Dim n As Long
 
 t_tr.ForeColor = &H0
 If InStr(lst_tr.Text, "- UN") > 0 Then 'user selects 'un-assigned'
   trailer = Trim$(lst_tr.Text)
   notttr = True: t_tr = trailer: notttr = False
 Else 'build trailer info string
   trailer = Trim$(Mid$(lst_tr.Text, 3, 11))
   If Left$(trailer, 2) = "ST" Then
     trailer = Mid$(trailer, 3): n = Val(trailer)
     For j = 1 To lst_un.ListCount - 1
       If n = Left(lst_un.List(j), 8) Then
         lst_un.ListIndex = j: lst_un_DblClick: Exit For
       End If
     Next j
   Else
     dm2 = UCase$(Mid$(lst_tr, 16, 2)) & "' "
     If dm2 = "00" Then dm2 = ""
     If InStr(lst_tr.Text, "FLAT") Then
       dm = "FLAT "
     ElseIf InStr(lst_tr.Text, "BARN") Then
       dm = "BARN "
     ElseIf InStr(lst_tr.Text, "ROLL") Then
       dm = "RLUP "
     ElseIf InStr(lst_tr.Text, "HIGHWAY") Then
       dm = "HWY "
     End If
     If InStr(UCase$(lst_tr.Text), "LOGIST") > 0 Then dm = dm & "LGST "
     If InStr(UCase$(lst_tr.Text), "TAIL") > 0 Then dm = dm & "TG "
     If InStr(UCase$(lst_tr.Text), "HEAT") > 0 Then dm = dm & "HEAT "
     If InStr(UCase$(lst_tr.Text), "GRAFFI") > 0 Then dm = dm & "GRAF "
     notttr = True: t_tr = BldTrlr: notttr = False
   End If
 End If
LTRDC:
 frtr.Visible = False: fr1.Enabled = True
 If t_tr <> "" Then
   dm = trailer: t_tro = "SP"
   If Not chksvcno Then ChkService
 End If
 If trlrapp.Visible Then trlrapp.SetFocus
End Sub
Private Function BldTrlr$()
 Dim dm As String, dm2 As String
 
 trailer = Trim$(Mid$(lst_tr.Text, 3, 11))
 dm2 = Mid$(lst_tr, 16, 2) & "' "
 If dm2 = "00" Then dm2 = ""
 If InStr(lst_tr.Text, "Flat") Then
   dm = "FLAT "
 ElseIf InStr(lst_tr.Text, "Barn") Then
   dm = "BARN "
 ElseIf InStr(lst_tr.Text, "Roll") Then
   dm = "RLUP "
 ElseIf InStr(lst_tr.Text, "Highway") Then
   dm = "HWY "
 End If
 If InStr(lst_tr.Text, "Logist") > 0 Then
   dm = dm & "LGST "
 ElseIf InStr(lst_tr.Text, "Tail") > 0 Then
   dm = dm & "  TG "
 End If
 BldTrlr = "SPDY " & dm2 & dm & trailer
End Function


Private Sub lst_tr_KeyDown(KeyCode%, Shf%)
 Select Case KeyCode
   Case vbKeyEscape: c_canctr_Click
   Case vbKeyReturn, vbKeyExecute: lst_tr_DblClick
 End Select
End Sub
Private Sub lst_tr_KeyPress(KeyAscii%)
 Select Case KeyAscii
   Case 48 To 57, 65 To 90
   Case 97 To 122: KeyAscii = KeyAscii - 32
   Case Else: KeyAscii = 0
 End Select
 If KeyAscii > 0 Then
   t_trfr = Chr$(KeyAscii)  'user keypress on list - write char to list-find field
   If t_trfr.Visible Then t_trfr.SetFocus
 End If
End Sub

Private Sub o_min_Click()
 o_min.Value = False
 WindowState = 1
End Sub

Private Sub mdte_Change()
 dl = Format$(mdte, "DDDD")
End Sub

Private Sub init_KeyDown(KeyCode%, Shf%) 'Initials field function keu
 'NOTE: see initials validation & set def. terminal in "init_Validate"
 Select Case KeyCode
   Case vbKeyEscape: init = ""
   Case vbKeyDown, vbKeyReturn, vbKeyExecute: init_Validate False 'force initials validation on field exit
 End Select
End Sub
Private Sub mdte_KeyDown(KeyCode%, Shf%)
 Select Case KeyCode
   Case vbKeyReturn, vbKeyExecute: c_un_Click  '''unit.SetFocus
 End Select
End Sub
Private Sub descr_KeyDown(KeyCode%, Shf%)
 Select Case KeyCode
   Case vbKeyDown, vbKeyReturn, vbKeyExecute
     If orig.Visible Then orig.SetFocus
   Case vbKeyUp: ldsht.SetFocus
 End Select
End Sub
Private Sub dest_KeyDown(KeyCode%, Shf%)
 Select Case KeyCode
   Case vbKeyDown, vbKeyReturn, vbKeyExecute
     newrow = 0: curow = 0
     If g2.Row = 0 Then
       HiRo 0, 1, 1
     Else
       If oldrow > -1 Then HiRo oldrow, 0, 1
       HiRo 0, 1, 1
     End If
     oldrow = newrow
     If g2.Visible Then g2pro.SetFocus
   Case vbKeyUp
     If orig.Visible Then orig.SetFocus
   Case vbKeyPageDown, vbKeyPageUp: c_de_Click
 End Select
End Sub
Private Sub ldsht_KeyDown(KeyCode%, Shf%)
 Select Case KeyCode
   Case vbKeyDown, vbKeyReturn, vbKeyExecute
     If descr.Visible Then descr.SetFocus
   Case vbKeyUp: seal.SetFocus
 End Select
End Sub
Private Sub orig_KeyDown(KeyCode%, Shf%)
 Select Case KeyCode
   Case vbKeyDown, vbKeyReturn, vbKeyExecute: dest.SetFocus
   Case vbKeyUp: descr.SetFocus
   Case vbKeyPageDown, vbKeyPageUp: c_or_Click
 End Select
End Sub
Private Sub driver_KeyDown(KeyCode%, Shf%)
 Select Case KeyCode
   Case vbKeyDown, vbKeyReturn, vbKeyExecute
     If seal.Visible Then seal.SetFocus
   Case vbKeyUp
     If trlrapp.Visible Then trlrapp.SetFocus
 End Select
End Sub
Private Sub seal_KeyDown(KeyCode%, Shf%)
 Select Case KeyCode
   Case vbKeyDown, vbKeyReturn, vbKeyExecute
     If ldsht.Visible Then ldsht.SetFocus
   Case vbKeyUp
     If driver.Visible Then driver.SetFocus
 End Select
End Sub

Private Sub trip_KeyDown(KeyCode%, Shf%)
 Select Case KeyCode
   Case vbKeyEscape: trip = ""
   Case vbKeyReturn, vbKeyExecute
     tz$ = trip
     ClrFrm
     trip = tz$
     c_ldtrp_Click
   Case vbKeyV
     If Shf = 2 Then
       dm3$ = Trim$(Clipboard.GetText)
       If (Len(dm3$) = 6 And IsNumeric(dm3$) = True) Then trip = dm3$
     End If
 End Select
End Sub

Private Sub trlrapp_KeyDown(KeyCode%, Shf%)
 Select Case KeyCode
   Case vbKeyUp: If t_tr.Visible Then t_tr.SetFocus
   Case vbKeyDown, vbKeyReturn, vbKeyExecute: driver.SetFocus
 End Select
End Sub

Private Sub descr_KeyPress(KeyAscii%)
 KeyAscii = AlphaFilt(KeyAscii, False) 'key, allow extended chars
End Sub
Private Sub dest_KeyPress(KeyAscii%)
 KeyAscii = NumFilt(KeyAscii, True, dest.SelStart) 'key, prevent leading zero, cursor pos
End Sub
Private Sub init_KeyPress(KeyAscii%)
 Select Case KeyAscii
   Case 8, 65 To 90 'bkspace, A - Z
   Case 97 To 122: KeyAscii = KeyAscii - 32 'uppercase a - z
   Case Else: KeyAscii = 0 'all other keys filtered out
 End Select
End Sub
Private Sub ldsht_KeyPress(KeyAscii%)
 KeyAscii = AlphaFilt(KeyAscii, True)
End Sub
Private Sub orig_KeyPress(KeyAscii%)
 KeyAscii = NumFilt(KeyAscii, True, orig.SelStart)
End Sub
Private Sub driver_KeyPress(KeyAscii%)
 KeyAscii = AlphaFilt(KeyAscii, True)
End Sub
Private Sub seal_KeyPress(KeyAscii%)
 KeyAscii = AlphaFilt(KeyAscii, True)
End Sub
Private Sub trlrapp_KeyPress(KeyAscii%)
 Select Case KeyAscii
   Case 8, 50 To 55 'bkspace, 2 - 6
   Case Else: KeyAscii = 0  'filter out all other chars
 End Select
End Sub
Private Sub trip_KeyPress(KeyAscii%)
 Select Case KeyAscii
   Case 8, 48 To 57, 65 To 90 'bkspace, 0-9, A - Z
   Case 97 To 122: KeyAscii = KeyAscii - 32 'uppercase a - z
   Case Else: KeyAscii = 0 'all other keys filtered out
 End Select
End Sub

Private Sub trlrapp_GotFocus()
 g2profoc = False
 trlrapp.BackColor = &HFFFFFF
End Sub
Private Sub descr_GotFocus()
 g2profoc = False
 descr.BackColor = &HFFFFFF
End Sub
Private Sub dest_GotFocus()
 g2profoc = False
 dest.BackColor = &HFFFFFF
End Sub
Private Sub init_GotFocus()
 g2profoc = False
 init.BackColor = &HFFFFFF
End Sub
Private Sub ldsht_GotFocus()
 g2profoc = False
 ldsht.BackColor = &HFFFFFF
End Sub
Private Sub mdte_gotFocus()
 g2profoc = False
End Sub
Private Sub orig_GotFocus()
 g2profoc = False
 orig.BackColor = &HFFFFFF
End Sub
Private Sub driver_GotFocus()
 driver.BackColor = &HFFFFFF
End Sub
Private Sub driver_LostFocus()
 driver = Replace(driver, "'", ""): driver = Replace(driver, "#", ""): driver = Replace(driver, "\", ""): driver = Replace(driver, "/", "")
 driver.BackColor = &HD6E6E6
End Sub
Private Sub seal_GotFocus()
 g2profoc = False
 seal.BackColor = &HFFFFFF
End Sub
Private Sub trip_GotFocus()
 g2profoc = False
 trip.BackColor = &HFFFFFF
End Sub
Private Sub trip_LostFocus()
 trip.BackColor = &HD6E6E6
End Sub
Private Sub trlrapp_LostFocus()
 trlrapp.BackColor = &HD6E6E6
End Sub
Private Sub descr_LostFocus()
 descr.BackColor = &HD6E6E6
End Sub
Private Sub dest_LostFocus()
 If dest = "14" Then
   MsgBox vbCrLf & "                *********  PLEASE NOTE  *********" & vbCrLf & vbCrLf & _
                   " IF Trailer is Destined for GUILBAULT, USE Code 80." & vbCrLf & vbCrLf, , "Manifest"
 End If
 dest.BackColor = &HD6E6E6
End Sub
Private Sub init_LostFocus()
 If Len(init) <> 3 Then '#@#
   MsgBox "Please Enter Your Initials.", , "Manifest"
   init.SetFocus
   Exit Sub
 End If
 init.TabStop = False 'req'd to give 'pro' text field the focus by default
 init.BackColor = &HD6E6E6
End Sub
Private Sub ldsht_LostFocus()
 ldsht.BackColor = &HD6E6E6
End Sub
Private Sub orig_LostFocus()
 orig.BackColor = &HD6E6E6
 Select Case Val(orig) 'for Sameday (68) and terminal dock pickups - dest=orig
   Case 68, 37, 38, 69, 71, 72, 73: dest = orig
 End Select
 If Val(orig) = 14 Then
   orig = "": Exit Sub
 End If
End Sub
Private Sub seal_LostFocus()
 seal.BackColor = &HD6E6E6
End Sub

Private Sub init_Validate(Cancel As Boolean)
 'force user to enter valid 3-char initials per 'INI'
 If Not ucmdl And init = uinits Then Exit Sub
 If Not rsb.State = 0 Then rsb.Close
 uterm = "": urghts = "": uinits = "": utsvc = "": umail = "": gpissuer = "N/A"
 rsb.Open "select terminal,rights,mnfstzstat,settrailerservice,email,nam from uinits where init='" & init & "'"
 If Not rsb.EOF Then
   urghts = rsb(1): uinits = init: c_zstat.Visible = False: utsvc = rsb(3): umail = rsb(4): gpissuer = Trim$(rsb(5))
   If Len(gpissuer) < 2 Then gpissuer = "N/A"
   If Val(rsb(2)) = 1 Then c_zstat.Visible = True
   Select Case rsb(0)
     Case "AX", "LN", "MT", "TR", "WR", "UC": uterm = rsb(0)
   End Select
 End If
 rsb.Close
End Sub


Private Sub tmr_Timer() 'display data transfer to legacy via removal from list box
 Dim dm As String, dm2 As String, dm3 As String
 Dim i As Integer, j As Integer, k As Integer, n As Integer
 Dim dd As Double
 Dim b As Boolean
 
 l(46) = Format$(Now, "DDD DD-MMM-YYYY") 'update screen date (top-center)
 
 If forcetrailerlistrefresh Then
   j = 0
   Do
     If rsc.State = 0 Then Exit Do
     Pause 0.1
     j = j + 1
   Loop Until j = 10
   If intv < 14 Then intv = intv + 1
   GoTo TMRFRC
 End If
  
 If rsc.State = 1 Then Exit Sub  'recordset in use, wait for next timer event
 
 dd = DateValue(Now) + Format(TimeValue(Now), ".#0000") 'get current date-time in julian format
 If plst.ListCount = 0 Then GoTo TTMRST 'no pros in transfer list
 
 If l_ne = "NEW" And mdte <> Now Then mdte = Now
TTMRST:
 intv = intv + 1
 If intv = 15 Then  'every 30 sec. @ 2sec timer
   intv = 0
TMRFRC:
   n = lst_tr.ListIndex
   dm = "select trailer,owner,svccode from trlrsvc where owner='SZTG' order by trailer"
   rsc.Open dm
   lst_tr.Visible = False
   If rsc.EOF Then 'no service trailers, remove any previously marked
     rsc.Close
     For j = 0 To lst_tr.ListCount - 1
       If Left$(lst_tr.List(j), 2) <> "  " Then
         dm = Mid$(lst_tr.List(j), 3): lst_tr.RemoveItem j: lst_tr.AddItem "  " & dm, j
       End If
     Next j
   Else
     'get all service trailers into working array
     j = 0: Erase sv
     Do
       j = j + 1: ReDim Preserve sv(1 To j): sv(j).t = rsc(0): sv(j).o = rsc(1): sv(j).c = rsc(2)
       rsc.MoveNext
     Loop Until rsc.EOF
     rsc.Close
     For j = 0 To lst_tr.ListCount - 1
       dm3 = Trim$(Mid$(lst_tr.List(j), 3, 11)): dm2 = Left$(lst_tr.List(j), 2)
       For k = 1 To UBound(sv)
         If sv(k).t = dm3 Then
           dm = Mid$(lst_tr.List(j), 3): lst_tr.RemoveItem j: lst_tr.AddItem sv(k).c & dm, j: GoTo NSV
         End If
       Next k
       If dm2 <> "" Then
         dm = Mid$(lst_tr.List(j), 3): lst_tr.RemoveItem j: lst_tr.AddItem "  " & dm, j
       End If
NSV:
     Next j
   End If
   If n > -1 Then lst_tr.Selected(n) = True
   lst_tr.Visible = True
   
   If forcetrailerlistrefresh Then Exit Sub
   
 End If
 
 
 If tlst.ListCount = 0 Then Exit Sub
 b = False: dm = "select trans_ack_timestamp from trp_trans where trip = '"
TMR1:
 If tlst.ListCount > 0 Then
   For i = 0 To tlst.ListCount - 1
     j = InStr(tlst.List(i), Chr(9)) - 1
     If j > 0 Then
       rsc.Open dm & Left$(tlst.List(i), j) & "'"
       If Not rsc.EOF Then
         If rsc(0) <> "" Then
           rsc.Close: tlst.RemoveItem i: GoTo TMR1
         End If
       End If
       rsc.Close
     End If
   Next i
 End If
 If tlst.ListCount > 0 Then
   For i = 0 To tlst.ListCount - 1 'cycle thru un-transferred trips list
     If InStr(tlst.List(i), Chr(124)) = 0 Then
       j = InStr(tlst.List(i), Chr(9)) + 1
       If j > 1 Then
         If dd - Mid(tlst.List(i), j) > 0.0009 Then
           If b Then
             tlst.List(i) = tlst.List(i) & Chr(124)
           Else
             tlst.List(i) = tlst.List(i) & Chr(124)
             b = True: Exit For
           End If
         End If
       End If
     End If
   Next i
 End If
 plst.Refresh: tlst.Refresh
End Sub


'******* SUBROUTINES *********************************************************************

Private Sub DispTCntr()
 dm = "select count(*) from trip where date_modified='" & Format$(Now, "YYYY-MM-DD") & _
      "' and init='" & init & "'"
 rs.Open dm
 If rs.EOF Then pcntr = "0        " Else pcntr = rs(0) & "        "
 rs.Close
End Sub


'****** FUNCTIONS *************************************************************************
Private Function odOK(ct As Control, cl As Control) As Boolean
 Dim i As Integer 'check for entered-code in valid codes list
 For i = 0 To cl.ListCount - 1
   If Val(ct) = Val(Left$(cl.List(i), 2)) Then
     cl.ListIndex = i: ct.ForeColor = &H0 'valid code black
     odOK = True: Exit Function
   End If
 Next i
 ct.ForeColor = &HC0 'invalid code red, function returns False
End Function

Private Function MGetTempBlobImage(dm$, pr$, ityp$, fnam$) As Boolean 'index, pro, type, filename
 Dim dm2 As String, dm3 As String
 Dim ist As New ADODB.Stream
 Dim isc As New ADODB.Stream
 
 'schfr.Visible = True: schfr.Refresh
 On Error Resume Next
 Err = 0
 ist.Open: isc.Open              'init ADO stream objects as ubiquitous streams to start
 isc.Type = adTypeBinary         'set stream to binary (auto-converts text to bin data during 'copyto' op)
 'add a temporary indexed blob record for image from local image directory on server
 'NOTE: Unix-like path/filename syntax
 ipardir = "d"
 Select Case Val(ityp)
   Case 2 'STPC BOL image
     dm3 = Format$(dm, "000000000")
     dm2 = "',img=LOAD_FILE(""/" & ipardir & "/STPCIMG/"
     dm2 = "insert ignore into imgtemp set ndx='" & dm & dm2 & _
            Left$(dm3, 3) & "/" & Mid$(dm3, 4, 3) & "/" & dm3 & ".TIF"")"
   Case Else 'std. BOL
     dm2 = "',img=LOAD_FILE(""/" & ipardir & "/"
     dm2 = "insert ignore into imgtemp set ndx='" & dm & dm2 & _
            Left$(dm, 8) & ".SCF/00/" & Mid$(dm, 9, 2) & "/" & Mid$(dm, 11, 2) & _
           "/00" & Right$(dm, 6) & ".TIF"")"
 End Select
 'Clipboard.Clear: Clipboard.SetText dm2
 icmd.CommandText = dm2: Err = 0: icmd.Execute
 If Err <> 0 Then
   MsgBox "Image Not Found!", , "Manifest"
   MGetTempBlobImage = True: GoTo GTBIERR
 End If
 'retrieve blobbed image textstream into recordset
 ir.Open "select img FROM imgtemp where ndx='" & dm & "'"
 If Err <> 0 Then
   MsgBox "Error Retrieving Image (2): " & Err.Description, , "Manifest"
   icmd.CommandText = "delete from imgtemp where ndx='" & dm & "'"
   icmd.Execute: MGetTempBlobImage = True: GoTo GTBIERR
 End If
 If ir.EOF Then
   MsgBox "Error Retrieving Image (3 - Image Buffer EOF)", , "Manifest"
   icmd.CommandText = "delete from imgtemp where ndx='" & dm & "'"
   icmd.Execute: ir.Close: MGetTempBlobImage = True: GoTo GTBIERR
 End If
 ist.Flush
 ist.WriteText ir(0) 'clear & write blobimage text to ADO stream object
 If Err <> 0 Then
   MsgBox "Error Retrieving Image (4): " & Err.Description, , "Manifest"
   icmd.CommandText = "delete from imgtemp where ndx='" & dm & "'"
   icmd.Execute: ir.Close: MGetTempBlobImage = True: GoTo GTBIERR
 End If
 ir.Close
 'delete temporary blob record
 icmd.CommandText = "delete from imgtemp where ndx='" & dm & "'"
 icmd.Execute
 Err = 0
 ofs = ist.Size   'get stream 1 size
 ist.SetEOS       'set end-of-stream (file) marker
 ist.Position = 2 'advance stream pointer past first 2 chars to get rid of them
 isc.Flush
 ist.CopyTo isc, ofs - 2 'copy data to 2nd ADO stream object removing 1rst 2 chars & converting to binary
 ist.Flush               'clear stream 1
 If Err <> 0 Then
   MsgBox "Error Retrieving Image (5): " & Err.Description, , "Manifest"
   isc.Flush: MGetTempBlobImage = True: GoTo GTBIERR
 End If
 'write binary stream 2 data to local file
 If fnam = "" Then isc.SaveToFile sdir & "mprntmp\" & pr & ".tif", adSaveCreateOverWrite Else isc.SaveToFile fnam, adSaveCreateOverWrite
 ist.Close: isc.Close
 If Err <> 0 Then
   MsgBox "Error Retrieving Image (6): " & Err.Description, , "Manifest"
   MGetTempBlobImage = True
 End If
GTBIERR:
 Err = 0
 Set ist = Nothing: Set isc = Nothing
End Function


Private Sub PrntGPProBkGrnd(cpy%, oy#)
 'start new pro print page titles/bkgrnd lines
 'update 06Aug09 to print 3rdParty/Billto customer logo where configured
 
 Printer.FontName = "Arial": Printer.FontBold = True
 Printer.FontItalic = False: Printer.FontUnderline = False
   
 'print logo specific to pro
 Select Case cpy
   Case -1
      Printer.PaintPicture pic1, 0!, 0.015 + oy, 1.8, 0.554
 End Select
 
 'Holland aspect: 1.862
 'Ingram: 2.276
 'NewPenn: 2.549
 'ODFL: 1.0
 'Avery: 3.086
 
 
 'top-right corner copy no. .375 x .375
 Set pic1.Picture = LoadPicture("C:\Program Files\PROTRACE\trcrnrbk.JPG")
 Printer.PaintPicture pic1, 7.425, 0.14 + oy, 0.575, 0.36
 Set pic1.Picture = LoadPicture("C:\Program Files\PROTRACE\speedy.jpg")
  
 Printer.FontSize = 20
 Printer.ForeColor = &HFFFFFF
 cprnt 7.9, 0.125 + oy, "1"
 Printer.ForeColor = &H0
 
 'draw all horiz lines from top
 If oy > 4 Then 'bill copy separator line
   Printer.DrawWidth = 1: Printer.DrawStyle = vbDot
   Printer.Line (0, oy - 0.0625)-Step(8#, 0)
   Printer.DrawStyle = 0: Printer.DrawWidth = 3
 End If
 
 Printer.Line (0, 0 + oy)-Step(0, 0.0625) 'mark bill top corners, left
 Printer.Line (0, 0 + oy)-Step(0.0625, 0)
 Printer.Line (8#, 0 + oy)-Step(0, 0.0625) 'right
 Printer.Line (8#, 0 + oy)-Step(-0.0625, 0)
 
 Printer.Line (3.875, 0.1875 + oy)-Step(2.875, 0) 'bol, pro top
 Printer.Line (3.875, 0.5 + oy)-Step(2.875, 0)    'po# top
 Printer.Line (3.875, 0.8125 + oy)-Step(2.875, 0) 'bottom
 
 Printer.Line (1.90625, 0.375 + oy)-Step(1.96875, 0)  'orig, dest top
 Printer.Line (0, 0.6875 + oy)-Step(3.875, 0) 'consignee top
 
 Printer.Line (7, 0.5 + oy)-Step(1#, 0)    'COD amnt
 Printer.Line (7, 0.8125 + oy)-Step(1#, 0) 'mid
 Printer.Line (7, 1.125 + oy)-Step(1#, 0)  'bottom
 
 Printer.Line (0, 1.5 + oy)-Step(3.875, 0)  'shipper top
 Printer.Line (4.21875, 1.5 + oy)-Step(3.78125, 0) 'billto top
 
 Printer.Line (0, 2.375 + oy)-Step(8#, 0)   'top 2 lines of freight area
 Printer.Line (0, 2.6875 + oy)-Step(8#, 0)
 Printer.Line (0, 4.1875 + oy)-Step(8#, 0)  'bottom freight area
 Printer.Line (0, 4.5 + oy)-Step(8#, 0)
 Printer.Line (0, 4.8125 + oy)-Step(2.75, 0) 'left bottom
 Printer.Line (0, 5.125 + oy)-Step(2.75, 0)  'left bottom
 
 Printer.Line (5.625, 4.8125 + oy)-Step(0.59375, 0) 'in (bottom)
 Printer.Line (5.625, 5.0625 + oy)-Step(0.59375, 0) 'out (bottom)
  
 Printer.Line (6.59375, 4.8125 + oy)-Step(1.4375, 0) 'driver
 Printer.Line (6.75, 5.0625 + oy)-Step(1.25, 0) 'date del'd
 
 Printer.Line (3.015625, 4.71875 + oy)-Step(2.609375, 0) 'firm
 Printer.Line (3.3125, 4.9375 + oy)-Step(2.3125, 0)   'printname
 Printer.Line (3.28125, 5.15625 + oy)-Step(2.34375, 0) 'signature
 
 'all verticals from left ------------------------
 Printer.Line (0, 0.6875 + oy)-Step(0, 1.25)      'consignee
 Printer.Line (4.21875, 1.5 + oy)-Step(0, 0.4375) 'third party 'bill chrgs to
 
 Printer.Line (3.875, 0.1875 + oy)-Step(0, 0.625) 'bol, pkup unit left
 Printer.Line (5.5625, 0.1875 + oy)-Step(0, 0.625) 'mid line
 Printer.Line (6.75, 0.1875 + oy)-Step(0, 0.625) 'pro,date right
 
 Printer.Line (7#, 0.5 + oy)-Step(0, 0.625)  'COD
 Printer.Line (7, 0.5078125 + oy)-Step(1#, 0) 'top xtra line
 Printer.Line (7.0078125, 0.5 + oy)-Step(0, 0.3125) 'left xtra
 Printer.Line (7, 0.8046875 + oy)-Step(1#, 0) 'bottom xtra
 
 Printer.Line (1.90625, 0.375 + oy)-Step(0, 0.3125) 'orig
 Printer.Line (2.25, 0.375 + oy)-Step(0, 0.3125)    'dest
 Printer.Line (2.65625, 0.375 + oy)-Step(0, 0.3125) 'QST
 Printer.Line (2.96875, 0.375 + oy)-Step(0, 0.3125)  'P/C/T/N
 Printer.Line (3.375, 0.375 + oy)-Step(0, 0.3125)   'biller
  
 Printer.Line (0.5, 4.1875 + oy)-Step(0, 0.3125)   'req date
 Printer.Line (1.375, 4.1875 + oy)-Step(0, 0.3125)  'time
 Printer.Line (2.25, 4.1875 + oy)-Step(0, 0.3125) 'svc code
 Printer.Line (2.75, 4.1875 + oy)-Step(0, 0.9375)    'rev'd in good ..., firm, printname signature
 Printer.Line (0.625, 4.5 + oy)-Step(0, 0.625) 'trailer/bay
 Printer.Line (1.5, 4.5 + oy)-Step(0, 0.625)   'pieces/pieces
 Printer.Line (2.125, 4.5 + oy)-Step(0, 0.625)   'initials/initials
  
 Printer.Line (0.625, 2.6875 + oy)-(0.625, 4.1875 + oy) 'pcs
 Printer.Line (0.8125, 2.6875 + oy)-(0.8125, 4.1875 + oy) 'H/M
  
 Printer.Line (5.625, 4.1875 + oy)-Step(0, 0.875) 'right of Recv'd in good . . .
 Printer.Line (6.21875, 4.1875 + 0.3125 + oy)-Step(0, 0.96875 - 0.3125) 'left of driver, date del'd
 Printer.Line (6.625, 4.1875 + oy)-Step(0, 0.3125) 'right of CPC control, left of Out
 Printer.Line (7.078125, 4.1875 + oy)-Step(0, 0.3125) 'right of Out, left of Retn
 Printer.Line (7.53125, 4.1875 + oy)-Step(0, 0.3125) 'right of Retn, left of Clsd
 
 Printer.Line (0.90625, 2.375 + oy)-Step(0, 0.3125) 'route#
 Printer.Line (1.5, 2.375 + oy)-Step(0, 0.3125)    'byd scac
 Printer.Line (2#, 2.375 + oy)-Step(0, 0.3125)     'beyond rev
 Printer.Line (2.9375, 2.375 + oy)-Step(0, 0.3125) 'adv scac
 Printer.Line (3.5, 2.375 + oy)-Step(0, 0.3125)    'adv pro
 
 Printer.Line (4.8125, 2.375 + oy)-(4.8125, 4.1875 + oy)    'descr
 Printer.Line (5.40625, 2.6875 + oy)-(5.40625, 4.1875 + oy) 'class
 Printer.Line (5.6875, 2.375 + oy)-Step(0, 0.3125)         'weight, adv rev
 Printer.Line (6#, 2.6875 + oy)-(6#, 4.1875 + oy)           'rate
 Printer.Line (6.625, 2.375 + oy)-(6.625, 4.1875 + oy)      'chgs
 Printer.Line (7.53125, 2.375 + oy)-(7.53125, 4.5 + oy) 'ppd/col, CPC clsd
  
 Printer.Line (7.3125, 5.0625 + oy)-Step(0.125, -0.21875) 'date del'd separators
 Printer.Line (7.625, 5.0625 + oy)-Step(0.125, -0.21875)
 
 'title text
 Printer.FontSize = 7  'font=Arial 6pt Bold
 Printer.FontBold = True
 'main address
 Select Case cpy
   Case 0 'U-Can Express
     lprnt 3.875, 0.03125 + oy, "PHONE (416) 620-1717"
     lprnt 1.90625, 0.09375 + oy, "55 BROWNS LINE"
     lprnt 1.90625, 0.1975 + oy, "ETOBICOKE, ONTARIO  M8W 3S2"
   Case -1 'Speedy
     lprnt 3.875, 0.03125 + oy, "PHONE (416) 510-2034"
     lprnt 1.90625, 0.09375 + oy, "265 RUTHERFORD ROAD SOUTH"
     lprnt 1.90625, 0.1975 + oy, "BRAMPTON, ONTARIO  L6W 1V9"
 End Select
     
 Printer.FontSize = 6
 Printer.FontBold = False
 lprnt 0.03125, 0.6875 + oy, "CONSIGNEE"
 lprnt 0.03125, 1.5 + oy, "SHIPPER"
 lprnt 4.25, 1.5 + oy, "THIRD PARTY"  ''"BILL CHARGES TO"
 
 lprnt 1.9375, 0.375 + oy, "ORIG."
 lprnt 2.28125, 0.375 + oy, "DEST."
 lprnt 2.6875, 0.375 + oy, "QST"
 lprnt 3#, 0.375 + oy, "P/C/T/N"
 lprnt 3.40625, 0.375 + oy, "BILLER"
 
 lprnt 3.90625, 0.1875 + oy, "SHIPPER BOL NO."
 lprnt 3.90625, 0.5 + oy, "PICKUP UNIT"
 lprnt 5.59375, 0.1875 + oy, "PRO NUMBER"
 lprnt 5.59375, 0.5 + oy, "DATE"
 
 lprnt 7.21875, 0.5 + oy, "COD AMOUNT"
  
 lprnt 0, 2.375 + oy, "INBOUND TRAILER"
 lprnt 0.9375, 2.375 + oy, "ROUTE NO."
 lprnt 1.53125, 2.375 + oy, "BYD.SCAC"
 lprnt 2.03125, 2.375 + oy, "BEYOND REVENUE"
 lprnt 2.96875, 2.375 + oy, "ADV.SCAC"
 lprnt 3.53125, 2.375 + oy, "ADVANCE PRO"
 lprnt 4.84375, 2.375 + oy, "ADVANCE DATE"
 lprnt 5.71875, 2.375 + oy, "ADVANCE REVENUE"
 lprnt 6.65625, 2.375 + oy, "STG REVENUE"

 cprnt 0.3125, 2.6875 + oy, "PIECES"
 cprnt 0.625 + (0.8125 - 0.625) / 2#, 2.6875 + oy, "H/M"
 cprnt 0.8125 + (4.8125 - 0.8125) / 2#, 2.6875 + oy, "DESCRIPTION"
 cprnt 4.8125 + (5.40625 - 4.8125) / 2#, 2.6875 + oy, "CLASS"
 cprnt 5.40625 + (6# - 5.40625) / 2#, 2.6875 + oy, "WEIGHT(LBS)"
 cprnt 6# + (6.625 - 6#) / 2#, 2.6875 + oy, "RATE"
 cprnt 6.625 + (7.53125 - 6.625) / 2#, 2.6875 + oy, "CHARGES"
 lprnt 7.6, 2.6875 + oy, "PPD/COL"
  
 lprnt 0.03125, 4.1875 + oy, "APPT."
 lprnt 0.53125, 4.1875 + oy, "REQUESTED DATE"
 lprnt 1.40625, 4.1875 + oy, "TIME"
 lprnt 2.28125, 4.1875 + oy, "SVC.CODE"
  
 lprnt 5.65625, 4.1875 + oy, "CPC CONTROL NO."
 lprnt 6.65625, 4.1875 + oy, "OUT"
 lprnt 7.109375, 4.1875 + oy, "RETURN"
 lprnt 7.5625, 4.1875 + oy, "CPC CLSD"

 lprnt 0.03125, 4.5 + oy, "DATE"
 lprnt 0.65625, 4.5 + oy, "TRAILER/BAY"
 lprnt 1.53125, 4.5 + oy, "PIECES"
 lprnt 2.15625, 4.5 + oy, "INITIALS"
 lprnt 0.03125, 4.8125 + oy, "DATE"
 lprnt 0.65625, 4.8125 + oy, "TRAILER/BAY"
 lprnt 1.53125, 4.8125 + oy, "PIECES"
 lprnt 2.15625, 4.8125 + oy, "INITIALS"
 
 lprnt 2.78125, 4.640625 + oy, "FIRM"
 lprnt 2.78125, 4.859375 + oy, "PRINT NAME"
 lprnt 2.78125, 5.078125 + oy, "SIGNATURE"

 lprnt 5.65625, 4.5 + oy, "IN"
 lprnt 5.65625, 4.8125 + oy, "OUT"
 lprnt 6.25, 4.734375 + oy, "DRIVER"
 lprnt 6.25, 4.984375 + oy, "DATE DEL'D"
   
 'different font texts
 Printer.FontSize = 5.55
 lprnt 2.78125, 4.25 + oy, "RECEIVED IN GOOD CONDITION EXCEPT AS NOTED. WHERE APPLICABLE"
 lprnt 2.78125, 4.34375 + oy, "SHIPMENT DELIVERED WITH WRAP INTACT UNLESS OTHERWISE NOTED."
 Printer.FontSize = 6: Printer.Font = "Arial Narrow": Printer.FontBold = False
 lprnt 7.03125, 0.8125 + oy, "FREIGHT CHARGES DUE STG"
 Printer.Font = "Arial"
 Printer.FontSize = 5
 cprnt 7.531253, 4.8125 + oy, "DD" 'date del'd format
 lprnt 7.15625, 4.8125 + oy, "MM"
 lprnt 7.78125, 4.8125 + oy, "YY"
 lprnt 7.39, 5.078125 + oy, "REV.20151109DNY"
 
End Sub

Private Function GetSvcDateFromCons$() 'use points-guide zoning tables to determine service date
 'assumes probill recordset 'pb' available, already filtered to Speedy service area (cons province)
 Dim shzon As String, cnzon As String, cty As String
 Dim j As Integer, k As Integer, n As Integer
 Dim drdte As Date
 
 'service date requires shipment pickup date
 If IsDate(pb!Date) Then
   If Val(pb!Date) = 0 Then Exit Function
 Else
   Exit Function
 End If
 drdte = DateValue(pb!Date)
 
 'find shipper zone
 shzon = ""
 If usnorthbnd Then
   If drinport <> "" Then 'this will catch bulk of US northbound shipments
     Select Case Val(drinport) 'use Cdn port since port is determined via Speedy routing tables
        Case 395: shzon = "MONTREAL"
        Case 423: shzon = "LONDON"
        Case 453: shzon = "WINDSOR"
        Case 495: shzon = "TORONTO"
     End Select
   Else
     If entport = "351" Then shzon = "MONTREAL" 'anything coming thru LaColle, QC goes to Montreal
   End If
 Else 'domestic shipment since US southbound & W.Cda already filtered out
   'use origin (pickup) terminal
   Select Case pb!complete
     Case "T": shzon = "TORONTO"
     Case "M": shzon = "MONTREAL"
     Case "L": shzon = "LONDON"
     Case "W": shzon = "WINDSOR"
     Case "A": shzon = "KINGSTON"
   End Select
 End If
 If shzon = "" Then
   If Val(pb!ship_accnum) > 0 Then 'if a Speedy account then look for zone attribute
     rs.Open "select zone from sc where account_num='" & pb!ship_accnum & "'"
     If Not rs.EOF Then
       If rs(0) <> "" Then shzon = rs(0)
     End If
     rs.Close
   End If
 End If
 If shzon = "" Then
   'if here then have to try matching to shipper FSA/address
   If pb!ship_prov = "QC" Then cty = FixQCCity(pb!ship_city) Else cty = pb!ship_city
   If Len(pb!ship_postcode) > 2 Then
     rs.Open "select zone from can_ctys_zone_fsa where fsa='" & Left$(pb!ship_postcode, 3) & _
             "' group by zone order by zone"
     If Not rs.EOF Then 'FSA match, check if crosses zones
       rs.MoveNext
       If Not rs.EOF Then 'crosses zones, refine with city
         rs.Close
         rs.Open "select zone from can_ctys_zone_fsa where fsa='" & Left$(pb!ship_postcode, 3) & _
                 "' and city='" & mySav(cty) & "' group by zone order by zone"
         If Not rs.EOF Then 'match with FSA + city
           rs.MoveNext
           If Not rs.EOF Then 'crosses zones! cannot be resolved with FSA+City
             rs.Close
           Else 'unique zone found with FSA + City
             rs.MoveFirst: shzon = rs(0): rs.Close
           End If
         Else 'no match found with FSA+City (probably misspelled city?)
           rs.Close: GoTo SHZONCYPV
         End If
       Else 'found unique zone by FSA only
         rs.MoveFirst: shzon = rs(0): rs.Close 'found zone by FSA alone
       End If
     Else 'no FSA match, try city/prov only
       rs.Close
SHZONCYPV:
       rs.Open "select zone from can_ctys_zone_fsa where city='" & mySav(cty) & _
               "' and prov='" & pb!ship_prov & "' group by zone order by zone"
       If Not rs.EOF Then
         rs.MoveNext
         If Not rs.EOF Then 'multiple zones match
           rs.Close 'shipzone cannot be resolved
         Else
           rs.MoveFirst: shzon = rs(0): rs.Close
         End If
       End If
     End If
   Else 'no shipper FSA, try shipper city/prov
     GoTo SHZONCYPV
   End If
 End If
 If rs.State <> 0 Then rs.Close
 
 If shzon = "" Then Exit Function 'cannot compute service date without shipper zone
 
 cnzon = "": cty = ""
 If Val(pb!cons_accnum) > 0 Then 'if a Speedy account then look for zone attribute
   rs.Open "select zone from sc where account_num='" & pb!cons_accnum & "'"
   If Not rs.EOF Then
     If rs(0) <> "" Then
       cnzon = rs(0): rs.Close: GoTo STSVCDATE
     End If
   End If
   rs.Close
 End If
 
 If Left$(pb!cons_city, 6) = "SAINT " Then cty = Replace(pb!cons_city, "SAINT ", "ST ", , 1) Else cty = pb!cons_city
 
 'if here then have to try matching to consignee FSA/address
 If Len(pb!cons_postcode) > 2 Then
   rs.Open "select zone from can_ctys_zone_fsa where fsa='" & Left$(pb!cons_postcode, 3) & _
           "' group by zone order by zone"
   If Not rs.EOF Then 'FSA match, check if crosses zones
     rs.MoveNext
     If Not rs.EOF Then 'crosses zones, refine with city
       rs.Close
       rs.Open "select zone from can_ctys_zone_fsa where fsa='" & Left$(pb!cons_postcode, 3) & _
               "' and city='" & mySav(cty) & "' group by zone order by zone"
       If Not rs.EOF Then 'match with FSA + city
         rs.MoveNext
         If Not rs.EOF Then 'crosses zones! cannot be resolved with FSA+City
           rs.Close
         Else 'unique zone found with FSA + City
           rs.MoveFirst: cnzon = rs(0): rs.Close
         End If
       Else 'no match found with FSA+City (probably misspelled city?)
         rs.Close: GoTo CNZONCYPV
       End If
     Else 'found unique zone by FSA only
       rs.MoveFirst: cnzon = rs(0): rs.Close 'found zone by FSA alone
     End If
   Else 'no FSA match, try city/prov only
     rs.Close
CNZONCYPV:
     rs.Open "select zone from can_ctys_zone_fsa where city='" & mySav(cty) & _
             "' and prov='" & pb!cons_prov & "' group by zone order by zone"
     If Not rs.EOF Then
       rs.MoveNext
       If Not rs.EOF Then 'multiple zones match
         rs.Close 'conszone cannot be resolved
       Else
         rs.MoveFirst: cnzon = rs(0): rs.Close
       End If
     End If
   End If
 Else 'no consignee FSA, try consignee city/prov
   GoTo CNZONCYPV
 End If
 If rs.State <> 0 Then rs.Close
 
 If cnzon = "" Then Exit Function 'cannot compute service date without consignee zone
  
STSVCDATE:
 
 'get service days between shipper & consignee zones
 rs.Open "select " & Replace(Trim$(shzon), " ", "_") & _
         " from can_ctys_zonematrix where zone='" & Trim$(cnzon) & "'"
 If rs.EOF Then
   rs.Close: Exit Function
 End If
 n = rs(0): rs.Close
 k = 0: j = 0 'init working service day counter & calendar day counter
 Do
   j = j + 1
   Select Case Weekday(drdte + j)
     Case 1, 7 'Sun, Sat - no increment service day count, fall-thru & increment day counter j next around
     Case Else
       rs.Open "select weekday from holidays where date='" & Format$(drdte + j, "YYYY-MM-DD") & "'"
       If rs.EOF Then k = k + 1 'not a holiday, increment service day count
       rs.Close
   End Select
 Loop Until k >= n Or j = 32 'count up thru calendar days until service days count = zone service matrix
 GetSvcDateFromCons = Format$(drdte + j, "YYYY-MM-DD")
 
End Function

Private Function FixQCCity$(cty$)
 Dim dm As String
 dm = Replace(cty, " ", "-")
 If Left$(dm, 3) = "SAINT-" Then dm = Replace(dm, "SAINT-", "ST-", , 1)
 If Left$(dm, 4) = "SAINTE-" Then dm = Replace(dm, "SAINTE ", "STE-", , 1)
 FixQCCity = dm
End Function

Private Function USFChkDig$(spro$)
 Dim j As Integer, tev As Integer, tod As Integer
 Dim p(1 To 10) As Integer
 If Len(spro) <> 10 Then Exit Function
 For j = 1 To 10: p(j) = Mid(spro, j, 1): Next j
 tod = 3 * (p(1) + p(3) + p(5) + p(7) + p(9)): tev = p(2) + p(4) + p(6) + p(8) + p(10)
 USFChkDig = CStr((tod + tev) Mod 10)
End Function

Private Function USF9ChkDig$(pro$)
 Dim b As Double, q As Double, r As Long
 Dim s() As String
 If Len(pro) <> 9 Then Exit Function
 b = Mid(pro, 4, 6): q = b / 11#: s = Split(CStr(q), ".")
 If UBound(s) <= 0 Then
   USF9ChkDig = "0"
 Else
   r = 11 - (b - (Fix(q) * 11))
   If r = 10 Then USF9ChkDig = "X" Else USF9ChkDig = CStr(r)
 End If
End Function


Private Function GetTempLogoImage(rec$, typ$) As Boolean 'retrieve blob logo image file to local file for printing | typ var added 16Aug13(DNY) #@#
 Dim ist As New ADODB.Stream
 Dim isc As New ADODB.Stream
 On Error Resume Next
 Err = 0: ist.Open: isc.Open  'init ADO stream objects as ubiquitous streams to start
 isc.Type = adTypeBinary      'set stream to binary (auto-converts text to bin data during 'copyto' op)
 ir.Open "select img from custlogo where recno='" & rec & "'" 'retrieve blob image textstream
 If Err <> 0 Then
   ir.Close: MsgBox "Error Retrieving Image (1): " & Err.Description: GetTempLogoImage = True: GoTo GTLIEND
 End If
 If ir.EOF Then
   ir.Close: MsgBox "Error Retrieving Image (2 - Image Buffer EOF)": GetTempLogoImage = True: GoTo GTLIEND
 End If
 ist.Flush: ist.WriteText ir(0) 'clear & write blobimage text to ADO stream object
 If Err <> 0 Then
   ir.Close: MsgBox "Error Retrieving Image (3): " & Err.Description: GetTempLogoImage = True: GoTo GTLIEND
 End If
 ir.Close: ofs = ist.Size   'get stream 1 size
 ist.SetEOS: ist.Position = 2 'set end-of-stream (file) marker, advance stream pointer past first 2 chars to get rid of them
 isc.Flush: ist.CopyTo isc, ofs - 2 'copy data to 2nd ADO stream object removing 1rst 2 chars & converting to binary
 ist.Flush        'clear stream 1
 If Err <> 0 Then
   MsgBox "Error Retrieving Image (4): " & Err.Description: isc.Flush: GetTempLogoImage = True: GoTo GTLIEND
 End If
 If typ = "l" Then '#@#
   Select Case logotype '#@#
     Case "gif": isc.SaveToFile sdir & "logoimg.gif", adSaveCreateOverWrite
     Case "jpg": isc.SaveToFile sdir & "logoimg.jpg", adSaveCreateOverWrite
   End Select
 ElseIf typ = "w" Then
   Select Case wlogotype '#@#
     Case "wgif": isc.SaveToFile sdir & "logowmimg.gif", adSaveCreateOverWrite
     Case "wjpg":  isc.SaveToFile sdir & "logowmimg.jpg", adSaveCreateOverWrite
   End Select
 End If
 ist.Close: isc.Close
 If Err <> 0 Then
   MsgBox "Error Retrieving Image (5): " & Err.Description: GetTempLogoImage = True: GoTo GTLIEND
 End If
GTLIEND:
 Set ist = Nothing: Set isc = Nothing
End Function


Private Sub Rectangle_barc0()
 Dim v(1 To 8) As POINTAPI, n As Long
 Dim dx As Integer, dy As Integer
 barc(0).Cls
 
 v(1).x = 0: v(1).y = 0
 v(2).x = barc(0).ScaleWidth: v(2).y = v(1).y
 v(3).x = v(2).x: v(3).y = barc(0).ScaleHeight
 v(4).x = v(1).x: v(4).y = v(3).y
 barc(0).ForeColor = &H0: barc(0).FillColor = &H0: barc(0).FillStyle = 0
 n = Polygon(barc(0).hdc, v(1), 4) 'draw rectangle
 
 dx = 0.05 * barc(0).ScaleWidth: dy = 0.05 * barc(0).ScaleHeight
 v(1).x = dx: v(1).y = dy
 v(2).x = barc(0).ScaleWidth - dx: v(2).y = v(1).y
 v(3).x = v(2).x: v(3).y = barc(0).ScaleHeight - dy
 v(4).x = v(1).x: v(4).y = v(3).y
 barc(0).ForeColor = &H606060: barc(0).FillColor = &HFFFFFF: barc(0).FillStyle = 0
 n = Polygon(barc(0).hdc, v(1), 4) 'draw rectangle
 dx = 0.075 * barc(0).ScaleWidth: dy = 0.081 * barc(0).ScaleHeight 'note correction for wid > ht
 v(1).x = dx: v(1).y = dy
 v(2).x = barc(0).ScaleWidth - dx: v(2).y = v(1).y
 v(3).x = v(2).x: v(3).y = barc(0).ScaleHeight - dy
 v(4).x = v(1).x: v(4).y = v(3).y
 barc(0).ForeColor = &HC0C0C0
 n = Polygon(barc(0).hdc, v(1), 4) 'draw rectangle
 barc(0).FillColor = &H0      'draw horiz separator line
 barc(0).Line (dx, barc(0).ScaleHeight / 3.5)-(barc(0).ScaleWidth - dx, barc(0).ScaleHeight / 3.5), &HC0C0C0
 barc(0).Refresh: barc(0).ForeColor = &H0: barc(0).FillStyle = 0
End Sub

Private Sub Octagon_barc0() 'use Windows Graphics Design Interface dll (gdi32.dll) Polygon function
 Dim v(1 To 8) As POINTAPI, n As Long 'see Global for type var descrip
 Dim dx As Integer, dy As Integer
 barc(0).Cls
 
 v(1).x = 0: v(1).y = 0.3 * barc(0).ScaleHeight
 v(2).x = 0.3 * barc(0).ScaleWidth: v(2).y = 0
 v(3).x = 0.7 * barc(0).ScaleWidth: v(3).y = v(2).y
 v(4).x = barc(0).ScaleWidth: v(4).y = v(1).y
 v(5).x = v(4).x: v(5).y = 0.7 * barc(0).ScaleHeight
 v(6).x = v(3).x: v(6).y = barc(0).ScaleHeight
 v(7).x = v(2).x: v(7).y = v(6).y
 v(8).x = v(1).x: v(8).y = v(5).y
 barc(0).ForeColor = &H0: barc(0).FillColor = &H0: barc(0).FillStyle = 0
 n = Polygon(barc(0).hdc, v(1), 8): barc(0).Refresh
 
 dx = 0.05 * barc(0).ScaleWidth: dy = 0.05 * barc(0).ScaleHeight 'set white outside boundary widths (squeeze shape into picture box)
           'on 7/8" square picturebox (1225 twips), '.05' creates a 0.75" square octagon with 1/16" white boundary on all sides
 'set 8 vertices of octagon to fit into picturebox
 v(1).x = dx: v(1).y = 0.3 * barc(0).ScaleHeight
 v(2).x = 0.3 * barc(0).ScaleWidth: v(2).y = dy
 v(3).x = 0.7 * barc(0).ScaleWidth: v(3).y = v(2).y
 v(4).x = barc(0).ScaleWidth - dx: v(4).y = v(1).y
 v(5).x = v(4).x: v(5).y = 0.7 * barc(0).ScaleHeight
 v(6).x = v(3).x: v(6).y = barc(0).ScaleHeight - dy
 v(7).x = v(2).x: v(7).y = v(6).y
 v(8).x = v(1).x: v(8).y = v(5).y
 barc(0).ForeColor = &H606060: barc(0).FillColor = &HFFFFFF '75% black 1-pix line, white fill
 barc(0).FillStyle = 0        'solid fill with diagonal crosshatch instead of solid color
 n = Polygon(barc(0).hdc, v(1), 8) 'draw octagon
 barc(0).Refresh
 dx = 0.081 * barc(0).ScaleWidth: dy = 0.08 * barc(0).ScaleHeight
 v(1).x = dx: v(1).y = 0.3 * barc(0).ScaleHeight
 v(2).x = 0.3 * barc(0).ScaleWidth: v(2).y = dy
 v(3).x = 0.7 * barc(0).ScaleWidth: v(3).y = v(2).y
 v(4).x = barc(0).ScaleWidth - dx: v(4).y = v(1).y
 v(5).x = v(4).x: v(5).y = 0.7 * barc(0).ScaleHeight
 v(6).x = v(3).x: v(6).y = barc(0).ScaleHeight - dy
 v(7).x = v(2).x: v(7).y = v(6).y
 v(8).x = v(1).x: v(8).y = v(5).y
 barc(0).ForeColor = &HB0B0B0: barc(0).FillColor = &HD0D0D0    'light gray line & fill color
 barc(0).FillStyle = 5 '': barc(0).DrawStyle = 2 'pattern fill with diagonal crosshatch
 n = Polygon(barc(0).hdc, v(1), 8) 'draw octagon
 barc(0).ForeColor = &H0: barc(0).FillStyle = 0: barc(0).DrawStyle = 0: barc(0).FillColor = &H0 'black
End Sub

Private Sub Clean_pgi()
 Dim j As Integer, n As Integer, k As Integer
 n = UBound(pgi)
 For j = 0 To n 'blank un-wanted array entries
   If Len(Trim$(pgi(j))) = 1 Then pgi(j) = ""
   If InStr(pgi(j), "=E=") > 0 Then pgi(j) = ""
   If Trim$(pgi(j)) <> "" And Len(pgi(j)) < 13 Then pgi(j) = pgi(j) & Space(13 - Len(pgi(j)))
 Next j
RECLNPGI:
 n = UBound(pgi)
 If n = 0 Then Exit Sub
 If Trim$(pgi(n)) = "" Then
   ReDim Preserve pgi(n - 1): GoTo RECLNPGI 'remove last index, re-check array
 End If
 For j = 0 To n
   If Trim$(pgi(j)) = "" Then  'blank entry left from the first cleaning step
     For k = j To n - 1 'move all indices up by 1
       pgi(k) = pgi(k + 1)
     Next k
     ReDim Preserve pgi(n - 1) 'remove last index
     GoTo RECLNPGI 're-check array
   End If
 Next j
 'if here then no blank lines remain
End Sub

Private Sub c_arr_Click()
 Exit Sub
End Sub

Private Sub c_prntdr_Click()
 Dim dm As String, dm2 As String, idno As String, spro As String
 Dim i As Integer
 Dim pwhse As Boolean
 
 gi = -999 '=print all shipments
 prndr.Show 1
 If gi = -999 Then Exit Sub
 Refresh
 
 'return vars
  'gs1 = SQL code to add manifest print job to history table
   'gs2 = door ** this is added to manifest record only for now **
    'gs3 = orig/whse copy ie. 01, 10 or 11 (one or both)
     'gs4 = reason
      'gi  = count of previous manifest-print events
      
 If rs.State <> 0 Then rs.Close
 pwhse = False
 On Error Resume Next
 
  
 '-- update door in parsmnfst record if entered
 If gs2 <> "" Then
   cmd.CommandText = "update parsmnfst set door='" & gs2 & "' where recno='" & main!rec & "'"
   Err = 0: cmd.Execute: Err = 0
 End If
 
 drmnfst = True
 
REDRPRN:

 For i = 0 To 206 Step 2  'cycle thru trip pros - successive pages start at 'i'
 
   If g2.TextMatrix(i, 17) = "" Then GoTo NXTPDPRO
   
   spro = g2.TextMatrix(i, 17): edipro = ""
     
   'print DR - whse/orig/both copies
   dm2 = ""
   Select Case gs3
     Case "10" 'orig only
       idno = PrntDR(spro, "ORIG", "M") 'errors displayed in function
       If idno = "FATAL" Then Exit For
       dm2 = ", orig='1'"
     Case "01" 'whse only
       idno = PrntDR(spro, "WHSE", "M") 'errors displayed in function
       If idno = "FATAL" Then Exit For
       dm2 = ", whse='1'"
     Case "11" 'both - print orig first then set indicator to whse and print again with manifest as separator
       idno = PrntDR(spro, "ORIG", "M") 'errors displayed in function
       If idno = "FATAL" Then Exit For
       ''idno = PrntDR(spro, "WHSE") 'errors displayed in function
       ''If idno = "FATAL" Then Exit For
       dm2 = ", orig='1', whse='1'"
   End Select
   
  ''DEBUG   Exit Sub  '//////////////\\\\\\\\\\\\\\\\\\\\\\\\\\///////////////////\\\\\\\\\\\\\\\\\
      
   If Not pwhse Then
     'add pro print record to history
     dm = "insert ignore into drprnhist set date='" & Format$(Now, "YYYY-MM-DD") & "', time='" & _
           Format$(Now, "Hh:Nn:Ss") & "', tripnum='" & Right$(trip, 6) & "', init='" & init & _
           "', pro='" & spro & "', edipro='" & edipro & "', idno='" & idno & "', src='r'"
     If gi > 0 Then dm = dm & ", reason='" & mySav(gs4) & "'"
     If dm2 <> "" Then dm = dm & dm2
     cmd.CommandText = dm: Err = 0: cmd.Execute
     If Err <> 0 Then
       dm2 = "ERROR: Fail Write PrintHistory Record! Aborting Print Job . . ." & vbCrLf & Err.Description
       MsgBox dm2, , "DR Print": Exit For
     End If
   End If
    
NXTPDPRO:
 Next i
  
 If gs3 = 11 Then 'orig+whse - print whse copies
   gs3 = "01": pwhse = True: GoTo REDRPRN
 End If
  
 'add manifest print record to history using gs1
 If Left$(gs1, 39) = "insert ignore into drprnhist set date='" Then
   cmd.CommandText = gs1: Err = 0: cmd.Execute
   If Err <> 0 Then
     dm2 = "WARNING: Fail Write Stack Manifest Print DR History Record!" & vbCrLf & Err.Description
     MsgBox dm2, , "DR Print"
   End If
 End If
 drmnfst = False

End Sub

Private Function PrntDR$(pro$, typ$, src$)
 'pro = Speedy pro#, typ = Orig, Whse copy
 Dim k As Integer, drcpy As Integer
 Dim dm As String, dm2 As String, dm3 As String, hdte As String, chkdig As String
 Dim splt() As String
 Dim msgasship As Boolean
 ''' already defined -> Dim rsb As New ADODB.Recordset  'Ops general use in addition to global 'rs' Ops recordset
 
 drcpy = 0
STPDR:
 drcpy = drcpy + 1
 If typ = "ORIG" And drcpy > Val(gs6) Then GoTo PRNTDREND
    
 '1.  gather all required data for print job

 'get probill record
 pb.Open "select * from probill where pronumber='" & pro & "'"
 If pb.EOF Then
   pb.Close
   MsgBox "Probill Record for " & pro & " NOT FOUND! Aborting Print Job . . ."
   PrntDR = "FATAL": GoTo PRNTDREND
 End If
 
 On Error Resume Next
  
 '-- check for re-print by type, build display string
 reprn = ""
 dm = "select count(*) from drprnhist where pro='" & pro & "' and cast(idno as unsigned) > '0' and "
 Select Case typ
   Case "ORIG": dm = dm & "orig='1'"
   Case "WHSE": dm = dm & "whse='1'"
 End Select
 
 '''rsb.ActiveConnection = cmd.ActiveConnection 'connect to master server since will be used with a write op
 rsb.Open dm
 If Not rsb.EOF Then
   If Val(rsb(0)) > 0 Then reprn = "RE-PRINT " & CStr((rsb(0) + 1))
 End If
 rsb.Close
  
 '-- generate next-number print job ID#
 cmd.CommandText = "update tripcntr set drprntctrl = last_insert_id(drprntctrl + 1)"
 Err = 0: cmd.Execute
 If Err <> 0 Then
   MsgBox "FAIL Generate Next DR Print Control Number!" & vbCrLf & vbCrLf & Err.Description, , "DR Print"
   PrntDR = "FATAL": GoTo PRNTDREND
 End If
 rsb.Open "select last_insert_id()" 'complete mySQL method that ensures only this user gets this next-number
 If rsb.EOF Or Err <> 0 Then
   rsb.Close
   MsgBox "FAIL Retrieve Next-Number DR Print Control Index!" & vbCrLf & vbCrLf & Err.Description, , "DR Print"
   PrntDR = "FATAL": GoTo PRNTDREND
 End If
 pndx = rsb(0): rsb.Close 'next-number print ctrl# retrieved
 loadid = "" '05Feb14

 '''global rsb.ActiveConnection = conn 'connect to slave server since used for read ops only from here on

 '-- get probill additional info record - determines # of pages reqd
 notes = "": pgc = 1: frzbl = False
 rs.Open "select note from probnotes where pro='" & pro & "'"
 If Not rs.EOF Then notes = rs(0)
 rs.Close
 '30 lines of additional info per page 64 chars wide
 If Replace(Trim$(notes), ".", "") = "" Then
   pgc = 1: ReDim pgi(0 To 0)
 Else
   pgi = Split(notes, vbCrLf) 'split into array of lineitems
   Clean_pgi 'remove blank & un-wanted entries (eg. =E= edit initials)
   If UBound(pgi) < 29 Then 'Note: pgi array is zero-based
     pgc = 1
   Else '
     pgc = Fix(CSng(UBound(pgi) + 1) / 30!)
     If (UBound(pgi) + 1) Mod 30 > 0 Then pgc = pgc + 1
   End If
 End If
 'freezable freight
 If UBound(pgi) > -1 Then 'check notes for freezable freight indication
   For i = 0 To UBound(pgi)
     dm = " " & Replace(pgi(i), "-", "") & " ": dm = Replace(dm, "*", ""):
     dm = Replace(dm, "#", ""): dm = Replace(dm, "@", "")
     If InStr(dm, " FREEZ") > 0 Then
       frzbl = True
     ElseIf InStr(dm, " HEAT ") > 0 Then
       frzbl = True
     End If
   Next i
 End If
 If Not frzbl Then 'check 'Remark' field
   dm = " " & pb(33) & " "
   If InStr(dm, " FREEZ") > 0 Then
     frzbl = True
   ElseIf InStr(dm, " HEAT ") > 0 Then
     frzbl = True
   End If
   If Not frzbl Then 'check 'Description' field
     dm = " " & pb(21) & " "
     If InStr(dm, " FREEZ") > 0 Then
       frzbl = True
     ElseIf InStr(dm, " HEAT ") > 0 Then
       frzbl = True
     End If
   End If
 End If
  
 'init US northbound customs vars
 csa4 = False: csac = False: cbsac = False: drinbond = False: transn = ""
 transn = "": brk = "": cdte = "": ctim = "": entport = "": cport = "": indte = ""
 drtrailer = "": drrun = "": drinport = "": drdoor = "": drcsano = ""
   
'check for US partner special features via Speedy pro ranges
 prnodfl = False: drscac = "": dr_tel = "": chkdig = "": prnward = False: prnsefl = False: prnexla = False: prnhmes = False: prnpyle = False
 Select Case Left$(pro, 4)
   Case "EXLA"
     drscac = "EXLA": dr_tel = "201679": prnexla = True: edipro = pb!edi_unit
   Case "HMES"
     drscac = "HMES": dr_tel = "4205": prnhmes = True
     If IsNumeric(pb!edi_unit) = True And Len(pb!edi_unit) = 9 Then '9Dig
       chkdig = USF9ChkDig(pb!edi_unit): edipro = pb!edi_unit       '9Dig
     ElseIf IsNumeric(pb!edi_unit) = True And Len(pb!edi_unit) = 10 Then
       chkdig = USFChkDig(pb!edi_unit): edipro = pb!edi_unit
     Else
       If Len(pro) = 13 Then '9Dig
         edipro = Right$(pro, 9): chkdig = USF9ChkDig(edipro) '9Dig
       Else
         edipro = Right$(pro, 10): chkdig = USFChkDig(edipro)
       End If
     End If
   Case "PYLE"
     drscac = "PYLE": dr_tel = "513183": prnpyle = True: edipro = pb!edi_unit
     If edipro = "" Then edipro = Right$(pro, 9)
   Case "SEFL"
     drscac = "SEFL": dr_tel = "399351": prnsefl = True: edipro = pb!edi_unit
     If Len(edipro) = 8 Then
        edipro = Right$(pro, 8): edipro = edipro & SEFLChkDig(edipro)
     End If
   Case "WARD"
     drscac = "WARD": dr_tel = "196031": prnward = True: edipro = pb!edi_unit
     If edipro = "" Then edipro = Right$(pro, 10)
   Case Else
     Select Case Val(pro)
       Case 10000000 To 99999999
         drscac = "NPME": dr_tel = "114178"
       Case 100000000 To 499999999, 500100000 To 999999999
         drscac = "RDWY": dr_tel = "65502"
       Case 1000000000 To 4999999999#
         drscac = "HMES": dr_tel = "4205": chkdig = USFChkDig(pro) 'Mod 10 USF check digit required for POD's
       Case 9500000000# To 9599999999#
         drscac = "SEFL": dr_tel = "399351": prnsefl = True: edipro = pb!edi_unit
         If Len(edipro) <> 9 Then
           edipro = Right$(pro, 8): edipro = edipro & SEFLChkDig(edipro)
         End If
       Case 9210000000# To 9210099999#
         drscac = "WARD": dr_tel = "196031": prnward = True: edipro = pb!edi_unit
       Case Else          '%^
         drscac = "SZTG"  '%^
     End Select
 End Select
 'check for partner special features via billto(s)
 Select Case pb!billed_account
   Case 196031
     drscac = "WARD": dr_tel = "196031": prnward = True: edipro = pb!edi_unit
     If edipro = "" Then edipro = Right$(pro, 10)
 End Select
   
 '-- first run at routing bill
 usnorthbnd = False: western = False: atlantic = False: ontque = False
 drbyndscac = "": drdlvyterm = "": drrow = ""
 
 Select Case pb!cons_prov
   Case "AB", "BC", "MB", "NB", "NF", "NL", "NS", "NT", "NU", "ON", "PE", "PQ", "QC", "SK", "YT"
     Select Case pb!ship_prov
       Case "AB", "BC", "MB", "NB", "NF", "NL", "NS", "NT", "NU", "ON", "PE", "PQ", "QC", "SK", "YT"
         usnorthbnd = False
       Case Else
         usnorthbnd = True
     End Select
   Case Else 'US destination, skip over northbound customs processing
     usnorthbnd = False
 End Select
 
 Select Case pb!cons_prov
   Case "AB", "BC", "MB", "NT", "NU", "SK", "YT"
     western = True: drdlvyterm = "WC"
     drrow = "KIDY-W": drbyndscac = "KIDY"
     Select Case Val(pb!billed_account)
       Case 457199, 512782, 512792: drrow = "KIDY-W": drbyndscac = "KIDY" '25Jun22  'drrow = "RSUT-82": drbyndscac = "RSUT"  '' 28May20 drrow = "West 82": drbyndscac = ""
     End Select
   Case "NB", "NS", "PE"
     ontque = True
   Case "NF", "NL"
     ontque = True
   Case "ON", "PQ", "QC"
     ontque = True
 End Select
  
 dm = "": dm2 = "": dm3 = ""
 drrowdc = "": drrowdct = "": drrowman = ""
  
 If Not usnorthbnd Then GoTo SKPPDRCSTMS
  
 '=== nortbound customs ============================
 
 'if US partner northbound then process customs info
 'get STPC processing record
 '                              0                               1                  2
 '                  3                4                5            6                7
 '                          8                          9
 rsb.Open "select cast(parspros.mnfstdatetime as char),parspros.portentry,parspros.clrtype, " & _
          "parspros.brkname,parspros.trailer,parspros.run,parspros.inport,parsmnfst.door, " & _
          "cast(parspros.indtetim as char),parspros.csarecno from parspros,parsmnfst " & _
          "where parspros.pro='" & pro & "' and parsmnfst.recno=parspros.mnfstrecno " & _
          "order by parspros.recno desc limit 1"
 If rsb.EOF Then '** should NEVER happen for northbound pro unless printed before declared Inland
   rsb.Close: GoTo PRNTDRCHKCSTMS
 End If
 splt = Split(rsb(0), " ") 'split date-time field - if CSA shipment then this can be 'clearance date-time'
 If UBound(splt) = 1 Then cdte = Format$(DateValue(splt(0)), "DD-MMM-YYYY") Else cdte = ""
 splt = Split(rsb(8), " ")
 If UBound(splt) = 1 Then indte = Format$(DateValue(splt(0)), "DD-MMM-YYYY") Else indte = "???"
 If Val(rsb(1)) > 0 Then entport = rsb(1) 'port of entry
 brk = Left$(Trim$(rsb(3)), 15)           'broker name
 If rsb(2) = "CSA" Then csac = True       'CSA shipment ** according to Stack Manifest only**
 drtrailer = rsb(4): drrun = rsb(5): drinport = rsb(6)
 If gs2 = "" Then drdoor = rsb(7) Else drdoor = gs2
 If Val(rsb(9)) > 0 Then drcsano = rsb(9)
 
 'now check customs (RNS) for CBSA clearance
PRNTDRCHKCSTMS:
 rs.Open "select trans_num,cast(date_time as char),port,code from customs where pro='" & pro & _
         "' and code in ('4','23') order by date_time limit 1"
 If Not rs.EOF Then
   If rs(0) = "CSA" Or rs(3) = 23 Then 'received CSA clearance from RNS!!
     csa4 = True  'indicate cleared CSA by RNS
   Else
     cbsac = True: transn = rs(0): cport = rs(2)
   End If
   splt = Split(rs(1), " ")
   If UBound(splt) = 1 Then 'if RNS clearance date then use over manifested date
     cdte = Format$(DateValue(splt(0)), "DD-MMM-YYYY"): ctim = Format$(TimeValue(splt(1)), "Hh:Nn")
   End If
 Else
   If Not csac Or Not csa4 Then drinbond = True
 End If
 rs.Close: rsb.Close
 If csa4 Then
   rsb.Open "select busno from parscsaimporters where recno='" & drcsano & "'"
   If Not rsb.EOF Then drcsano = rsb(0)
   rsb.Close
 End If
 
SKPPDRCSTMS:

 'check for logo to use on DR
 logorec = 20: dr_tel = "" 'defaults
 Select Case Left$(pro, 4)
   Case "EXLA": dr_tel = "(888) 588-0750": logorec = 50: GoTo DRLOGOXY
   Case "HMES": dr_tel = "(800) 535-2988": logorec = 62: GoTo DRLOGOXY  '13Feb2023(DN)
   Case "YELL": dr_tel = "(800) 535-2988": logorec = 62: GoTo DRLOGOXY
   Case "PYLE": dr_tel = "(800) 265-5351": logorec = 55: GoTo DRLOGOXY
   Case "SEFL": dr_tel = "(800) 265-5351": logorec = 37: GoTo DRLOGOXY
   Case "WARD": dr_tel = "(800) 265-5351": logorec = 31: GoTo DRLOGOXY
   Case Else
     If IsNumeric(pro) = True Then
       Select Case Val(pro)
          Case 100000000 To 999999999
             dr_tel = "(800) 535-2988": logorec = 62: GoTo DRLOGOXY
       End Select
     End If
 End Select
 'hard-coded customer
 Select Case pb!billed_account
   Case "512588", "513213", "513595", "513599"  '3M Canada by billto - GTA & Non-GTA"
     dr_tel = "(800) 265-5351": logorec = 52: GoTo DRLOGOXY
 End Select
 Select Case pb!billed_account
   Case "514208", "514209", "514210"  'Yellow PT, CA + US-13Feb2023(DNY)
     dr_tel = "(800) 535-2988": logorec = 62: GoTo DRLOGOXY
 End Select
 If pb!ship_accnum = "100297" Or (InStr(pb!ship_company, "TOSHIBA") > 0 And (Left$(pb_ship_address, 9) = "191 MCNAB" Or pb!ship_postcode = "L3R8H2")) Then
   dr_tel = "(800) 265-5351": logorec = 27: GoTo DRLOGOXY
 End If
 If pb!ship_accnum = "477502" Or (InStr(pb!ship_company, "B&C") > 0 And (Left$(pb_ship_address, 10) = "3389 STEEL" And Left$(pb!ship_city, 4) = "BRAM")) Then
   dr_tel = "(800) 265-5351": logorec = 28: GoTo DRLOGOXY
 End If
 If Left$(pb!ship_company, 5) = "ULINE" And pb!ship_prov = "ON" Then
   dr_tel = "(800) 265-5351": logorec = 36: GoTo DRLOGOXY
 End If
 If pb!billed_account = "377492" Then 'Mark's Work Warehouse
   dr_tel = "(800) 265-5351": logorec = 48: GoTo DRLOGOXY
 End If
 If pb!billed_account = "103890" Then 'Forbes-Hewlett Transportation
   dr_tel = "(800) 265-5351": logorec = 51: GoTo DRLOGOXY
 End If
 If pb!billed_account = "513465" Then 'Tarkett USA
   dr_tel = "(800) 265-5351": logorec = 56: GoTo DRLOGOXY
 End If
 
 'customer via billto link in account codes
   'if here then not partner logo or hard-coded customer logo, check billto for assigned logo
   If Val(pb!billed_account) > 0 Then 'check if billto has logo configured in its account code record
     rs.Open "select logorecnum,dr_copysperpage,dr_tel from scops where accno='" & pb!billed_account & _
             "' and logorecnum > '0'"
     If Not rs.EOF Then 'there is a logo reference
       If Trim$(rs(2)) = "" Then dr_tel = "  " Else dr_tel = rs(2)
       logorec = rs(0): rs.Close: GoTo DRLOGOXY 'buff logo image recno & partner tele, goto sizing check
     Else
       rs.Close: dr_tel = "": GoTo DRNOLOGOXY
     End If
     rs.Close
   End If
DRLOGOXY:
     'buff location/sizing params for logo
     ir.Open "select x,y,w,h,pspdy,imgtype from custlogo where recno='" & logorec & "'" 'imgtype field added 16Aug13(DNY) #@#
     If ir.EOF Then 'if no sizing record then go back to Speedy default
       ir.Close: dr_tel = ""
     Else
       logox = ir(0): logoy = ir(1): logow = ir(2): logoh = ir(3): logotype = ir(5) '#@#
       If ir(4) = "1" Then pspdy = True Else pspdy = False
       ir.Close
       If GetTempLogoImage(CStr(logorec), "l") Then dr_tel = "" 'get logo image into 'program file\stpc\stpcstackmanifest' dir
     End If
DRNOLOGOXY:
 If dr_tel = "" Then 'if get-cust-logo failed or no cust/partner logo then goto Speedy default
   logorec = 20 'new logo, was recno=13
   ir.Open "select x,y,w,h,pspdy,imgtype from custlogo where recno='20'"
   logox = ir(0): logoy = ir(1): logow = ir(2): logoh = ir(3): logotype = ir(5) '#@#
   If ir(4) = "1" Then pspdy = True Else pspdy = False
   ir.Close: b = GetTempLogoImage("20", "l"): dr_tel = "(800) 265-5351"
 End If
 'Holland aspect: 1.862, Ingram: 2.276, NewPenn: 2.549, ODFL: 1.0, Avery: 3.086

 '-- row assignment from routing and/or dock-check or accnt attribute
 If western Or atlantic Then
   'vars already set fall-thru
 ElseIf ontque And drrow = "" Then
    
   '--start--Amazon DC hard-coded routing 25Nov2021(DNY)
   If Left$(pb!cons_address, 9) = "12780 COL" And Left$(pb!cons_company, 2) = "AM" And Left$(pb!cons_postcode, 3) = "L7E" Then drrow = "DR49"
   Select Case pb!cons_accnum
           'DOI5      DOI6      DTO5
     Case "239692", "247886", "135180": drrow = "R46-47"
           'DON8      YYZ9
     Case "241103", "168028": drrow = "PIC"
     
     'all below updated 06May22
     Case "211272", "180586": drrow = "DR49" 'YYZ7
     Case "887950", "955343": drrow = "DR50" 'YYZ3
     Case "125099": drrow = "DR47"   'YYZ4
     Case "706031": drrow = "DR48"   'YYZ2
     Case "954948": drrow = "DR46"   'XCAC
     Case "219038": drrow = "DR52"   'YHM1  03Oct22(DNY)  was DR19
     Case "593554": drrow = "DR20"   'YYZ1
     Case "224700", "240911": drrow = "DR44" 'YOO1  25Jun22(DNY)
     
     Case "137968": drrow = "MIL"  '': GoTo DRROWDOCK 'Mark's Supply Kitchener
     
   End Select
   '--end--Amazon DC hard-coded
   
   '--start--Wayfair DC hard-coded routing 02Dec2021(DNY), updated 19Jul22(DNY)
   Select Case pb!cons_accnum
     Case "938306": drrow = "D53"
   End Select
   If drrow = "" And pb!cons_prov = "ON" And Left$(pb!cons_city, 6) = "MISSIS" And Left$(pb!cons_address, 8) = "2020 LOG" Then
     If InStr(pb!cons_company, "WAYFAIR") > 0 Then drrow = "D53"
   End If
   '--end--Wayfair DC hard-coded
   
   '--start-- SCM Mississauga hard-coded 19Jul22(DNY)
   Select Case pb!cons_accnum
     Case "73466", "909391": drrow = "R90-R91"
   End Select
   If drrow = "" And pb!cons_prov = "ON" And Left$(pb!cons_city, 6) = "MISSIS" And Left$(pb!cons_address, 9) = "6800 MARI" Then
     If InStr(pb!cons_company, "SCM") > 0 Or InStr(pb!cons_company, "MART") > 0 Then drrow = "R90-R91"
   End If
   '--end-- SCM Mississauga hard-coded 19Jul22(DNY)
 
 
   If drrow <> "" Then
     If drrow = "MIL" Then
       drdlvyterm = "ML": drbyndscac = "SZTG": GoTo DNY001
     Else
       drdlvyterm = "TR"
     End If
     drbyndscac = "SZTG": GoTo DRROWDOCK
   End If
 
   If Len(pb!cons_postcode) > 2 Then
     If uterm = "M" Then
       rs.Open "select dterm,row,scac from can_ctys_zone_fsa where fsa='" & _
                Left$(pb!cons_postcode, 3) & "' group by row order by row"
     Else
       rs.Open "select dterm,disp,scac,row from can_ctys_zone_fsa where fsa='" & _
                Left$(pb!cons_postcode, 3) & "' group by disp order by disp"
     End If
     If rs.EOF Then GoTo DRROWTRYCTY
     rs.MoveNext
     If rs.EOF Then
       rs.MoveFirst: drdlvyterm = rs(0): drrow = rs(1): drbyndscac = rs(2)
     Else
       rs.Close
       If uterm = "M" Then
         rs.Open "select dterm,row,scac,row from can_ctys_zone_fsa where fsa='" & _
                  Left$(pb!cons_postcode, 3) & "' and prov='" & pb!cons_prov & "' and city='" & _
                  mySav(pb!cons_city) & "' group by row order by row"
       Else
         rs.Open "select dterm,disp,scac,row from can_ctys_zone_fsa where fsa='" & _
                  Left$(pb!cons_postcode, 3) & "' and prov='" & pb!cons_prov & "' and city='" & _
                  mySav(pb!cons_city) & "' group by disp order by disp"
       End If
       If rs.EOF Then GoTo DRROWTRYCTY
       rs.MoveNext
       If rs.EOF Then
         rs.MoveFirst: drdlvyterm = rs(0): drrow = rs(1): drbyndscac = rs(2)
       End If
     End If
   Else
DRROWTRYCTY:
     If rs.State <> 0 Then rs.Close
     If uterm = "M" Then
       rs.Open "select dterm,row,scac,row from can_ctys_zone_fsa where prov='" & pb!cons_prov & _
               "' and city='" & mySav(pb!cons_city) & "' group by row order by row"
     Else
       rs.Open "select dterm,disp,scac,row from can_ctys_zone_fsa where prov='" & pb!cons_prov & _
               "' and city='" & mySav(pb!cons_city) & "' group by disp order by disp"
     End If
     If Not rs.EOF Then
       rs.MoveFirst
       If rs.EOF Then
         rs.MoveFirst: drdlvyterm = rs(0): drrow = rs(1): drbyndscac = rs(2)
       End If
     End If
   End If
 End If
 If rs.State <> 0 Then rs.Close

DRROWDOCK:

 '-- look for a dock-check record
 rs.Open "select terminal,row,trailer from dock_check where pro='" & pro & _
         "' order by date_time desc limit 1"
 If Not rs.EOF Then
 
    Select Case rs(0)
     Case "TR": dm = "TOR"
     Case "MT": dm = "MTL"
     Case "LN": dm = "LON"
     Case "AX": dm = "BRO"
     Case "BR": dm = "BRA"
     Case "WR": dm = "WIN"
     Case "PI": dm = "PIC"
     Case "MI": dm = "MIS"
     Case "ML": dm = "MIL"
     Case "VA": dm = "LAC" 'VAU
     Case "XX": dm = ""
     Case Else: dm = ""
   End Select
   
 
   If rs(1) <> "" Then
     drrowdc = "R-" & rs(1): drrowdct = dm ''rs(0)
     If dm <> "" Then dcrowf = dm & ": R-" & UCase$(rs(1))
   ElseIf rs(2) <> "" Then
     '12Jan21(DNY)
     Select Case rs(2)
       Case "EUGHOLD", "JOTIHOLD", "KENHOLD", "R4ANDRE"
         drrowdc = rs(2): drrowdct = rs(2): drrowf = rs(2): drrow = rs(2)
       Case Else
         If drrow = "KIDY-W" Then ''"RSUT-82" Then
           'fall-thru
         Else
           drrowdc = "T-" & UCase$(rs(2)): drrowdct = dm
           If dm <> "" Then dcrowf = dm & ": T-" & UCase$(rs(2))
         End If
     End Select
   End If
 End If
 rs.Close
   
DNY001:

 '-- req delivery, open/closed times, account codes, billto/3rd party
 reqdeldte = "": apptdate = DateValue("1900-01-01"): Erase bt
 '-- check for/build appointment detail display
 apptstr = ""
 If IsDate(pb!apptmnt_date) = True Then
   apptdate = DateValue(pb!apptmnt_date) '29May13
   If Val(pb!apptmnt_date) > 0 And Trim$(pb!door) <> "" Then
     reqdeldte = Format$(DateValue(pb!apptmnt_date), "YYYY-MM-DD")
     apptstr = Format$(DateValue(pb!apptmnt_date), "DDD DD-MMM-YYYY") 'appt date
     If Trim$(pb!door) <> "" Then                                     'From time
       dm = Trim$(pb!door)
       If Val(dm) = 0 Then dm = ""
       Select Case Len(dm)
         Case 3: dm = "0" & Left$(dm, 1) & ":" & Right$(dm, 2)
         Case 4: dm = Left$(dm, 2) & ":" & Right$(dm, 2)
       End Select
       apptstr = apptstr & " " & dm & Space(5 - Len(dm))
     Else
       apptstr = apptstr & Space(6)
     End If
     If Trim$(pb!apptmnt_time) <> "" Then                             'To time (if any)
       dm = Trim$(pb!apptmnt_time)
       If Val(dm) = 0 Then dm = ""
       Select Case Len(dm)
         Case 3: dm = "0" & Left$(dm, 1) & ":" & Right$(dm, 2)
         Case 4: dm = Left$(dm, 2) & ":" & Right$(dm, 2)
       End Select
       If dm = "" Then apptstr = apptstr & Space(6) Else apptstr = apptstr & "-" & dm & Space(5 - Len(dm))
     Else
       apptstr = apptstr & Space(6)
     End If
     If Trim$(pb!apptmnt_no) <> "" Then apptstr = apptstr & " " & pb!apptmnt_no 'appt#
   End If
 End If
  
'-- billto/3rd party
 If drscac = "SPDY" Then GoTo DOMSTHBT '%^

 If usnorthbnd And drscac <> "WARD" Then
   'US partner billto/3rd party
   If drscac <> "SEFL" Then
     rs.Open "select * from probusbt where pro='" & pro & "'"
     If Not rs.EOF Then
       If Val(rs(4)) > 0 Then bt(0) = rs(3) & rs(4)
       bt(1) = Left$(rs(5), 40): bt(2) = rs(6): bt(3) = rs(7): bt(4) = rs(8): bt(5) = rs(9)
     End If
     rs.Close
   End If
   'if not appt then check for req delivery date via US northbound EDI
   If reqdeldte = "" Then
     Select Case drscac
       Case "HMES"
         rs.Open "select cast(trailer as char) from probusf where pronumber='" & pro & "'"
         If Not rs.EOF Then
           If IsDate(rs(0)) = True And Val(rs(0)) > 0 Then reqdeldte = Format$(DateValue(rs(0)), "YYYY-MM-DD")
         End If
         rs.Close
       Case "EXLA"
         rs.Open "select cast(edidate as char) from probexla where pronumber='" & pro & "'"
         If Not rs.EOF Then
           If IsDate(rs(0)) = True And Val(rs(0)) > 0 Then reqdeldte = Format$(DateValue(rs(0)), "YYYY-MM-DD")
         End If
         rs.Close
       Case "PYLE"
         rs.Open "select cast(edidate as char) from probpyle where pronumber='" & pro & "'"
         If Not rs.EOF Then
           If IsDate(rs(0)) = True And Val(rs(0)) > 0 Then reqdeldte = Format$(DateValue(rs(0)), "YYYY-MM-DD")
         End If
         rs.Close
       Case "SEFL"
         rs.Open "select cast(edidate as char) from probsefl where pronumber='" & pro & "'"
         If Not rs.EOF Then
           If IsDate(rs(0)) = True And Val(rs(0)) > 0 Then reqdeldte = Format$(DateValue(rs(0)), "YYYY-MM-DD")
         End If
         rs.Close
       Case "WARD"
         rs.Open "select cast(edidate as char) from probward where pronumber='" & pro & "'"
         If Not rs.EOF Then
           If IsDate(rs(0)) = True And Val(rs(0)) > 0 Then reqdeldte = Format$(DateValue(rs(0)), "YYYY-MM-DD")
         End If
         rs.Close
     End Select
   End If
 Else
'%^
DOMSTHBT:
   'domestic/southbound billto/3rd party
   dm = "select account_num,company,shipaddress,city,prov,postcode from sc where "
   If Val(pb!billed_account) > 0 Then
     dm = dm & "account_num='" & pb!billed_account & "'"
   ElseIf Len(Trim$(pb!billto_company)) > 1 Then
     dm = dm & "company='" & myStr(pb!billto_company) & "'"
   End If
   rs.Open dm
   If Not rs.EOF Then
     If Val(pb!billed_account) > 0 Then bt(0) = rs(0)
     For k = 1 To 5: bt(k) = rs(k): Next k
   End If
   rs.Close
'-- 3PL LoadID/Confirmation No. --------------------
   dm = "select cn from procn where pro='" & pro & "' order by recno desc limit 1"
   If rs.State <> 0 Then rs.Close
   rs.Open dm
   If Not rs.EOF Then loadid = Trim$(rs(0))
   rs.Close
 End If

 '-- service date cleanup
 If reqdeldte = "" Then 'no service date - try Speedy service-guide tables vs. consignee address
   If drscac <> "" And Not usnorthbnd Then 'US southbound shipment
     'fall-thru
   Else 'all US northbound & domestic
     Select Case pb!cons_prov
       Case "NB", "NF", "NL", "NS", "ON", "PE", "QC" 'applies only to Speedy service area
         reqdeldte = GetSvcDateFromCons()
     End Select
   End If
 End If
  
 '-- process open/close times, account codes  12Nov13(DNY) shipper/consignee contact phone nos.
 Erase accs: Erase opcl: opcltim = "": acccnt = 0: rcvnghrs = False: apptreqd = False: Erase tel1: Erase tel2
 
 '\\\_/// Account Codes \\\_///--------------------------------------------------------------------------
 If Val(pb!ship_accnum) > 2 Then 'if shipper in Speedy accounts ...
   rs.Open "select appt_tel1, appt_tel2 from scops2 where accno='" & pb!ship_accnum & "'" 'check if shipper has contact phone nos.
   If Not rs.EOF Then
     If Len(rs(0)) > 6 Then tel1(0) = rs(0)
     If Len(rs(1)) > 6 Then tel2(0) = rs(1)
   End If
   rs.Close
 End If
 If Val(pb!cons_accnum) < 2 Then GoTo SKPDRACCS 'skip account codes if no match to Speedy cust master
 
 'receiving hrs on service date
 If reqdeldte <> "" Then
   dm = LCase$(Format$(DateValue(reqdeldte), "DDD")) 'get service date day-of-week & lookup rcvng hrs for that day
   dm = "select " & dm & ", " & dm & "_del_op, " & dm & "_del_cl, " & dm & "_delappt, appt_tel1, appt_tel2 " & _
        "from scops2 where accno='" & pb!cons_accnum & "'"
   If rs.State <> 0 Then rs.Close
   rs.Open dm
   If Not rs.EOF Then
     If Len(rs(4)) > 6 Then tel1(1) = rs(4) 'consignee contact phone nos.
     If Len(rs(5)) > 6 Then tel2(1) = rs(5)
     If Val(rs(0)) = 1 Then opcltim = "  CLOSED" Else opcltim = " " & rs(1) & "-" & rs(2)
     rs.Close
     '             0       1          2          3       4       5          6          7
     '             8       9         10         11      12      13         14         15
     '            16       17        18         19
     dm = "select mon,mon_del_op,mon_del_cl,mon_delappt,tue,tue_del_op,tue_del_cl,tue_delappt," & _
                 "wed,wed_del_op,wed_del_cl,wed_delappt,thu,thu_del_op,thu_del_cl,thu_delappt," & _
                 "fri,fri_del_op,fri_del_cl,fri_delappt from scops2 where accno='" & pb!cons_accnum & "'"
     rs.Open dm
     If Not rs.EOF Then
       If (Val(rs(3)) + Val(rs(7)) + Val(rs(11)) + Val(rs(15)) + Val(rs(19))) > 0 Then
         apptreqd = True: acccnt = acccnt + 1: accs(acccnt, 0) = "稟PPOINTMENT REQD": accs(accnt, 1) = "1" '1=type=fixed-coded, apptreqd-29May13
       End If
       If (Val(rs(1)) + Val(rs(5)) + Val(rs(9)) + Val(rs(13)) + Val(rs(17))) > 0 Then
         rcvnghrs = True
       ElseIf (Val(rs(0)) + Val(rs(4)) + Val(rs(8)) + Val(rs(12)) + Val(rs(16))) > 0 Then
         rcvnghrs = True
       End If
       If Val(rs(0)) = 1 Then
         opcl(1) = "M  Closed"
       Else
         If rcvnghrs Then
           opcl(1) = "M "
           If rs(1) <> "" And rs(2) <> "" Then opcl(1) = opcl(1) & rs(1) & "-" & rs(2)
         End If
       End If
       If Val(rs(4)) = 1 Then
         opcl(2) = "T  Closed"
       Else
         If rcvnghrs Then
           opcl(2) = "T "
           If rs(5) <> "" And rs(6) <> "" Then opcl(2) = opcl(2) & rs(5) & "-" & rs(6)
         End If
       End If
       If Val(rs(8)) = 1 Then
         opcl(3) = "W  Closed"
       Else
         If rcvnghrs Then
           opcl(3) = "W "
           If rs(9) <> "" And rs(10) <> "" Then opcl(3) = opcl(3) & rs(9) & "-" & rs(10)
         End If
       End If
       If Val(rs(12)) = 1 Then
         opcl(4) = "T  Closed"
       Else
         If rcvnghrs Then
           opcl(4) = "T "
           If rs(13) <> "" And rs(14) <> "" Then opcl(4) = opcl(4) & rs(13) & "-" & rs(14)
         End If
       End If
       If Val(rs(16)) = 1 Then
         opcl(5) = "F  Closed"
       Else
         If rcvnghrs Then
           opcl(5) = "F "
           If rs(17) <> "" And rs(18) <> "" Then opcl(5) = opcl(5) & rs(17) & "-" & rs(18)
         End If
       End If
     End If
   End If
   rs.Close
 End If
 
 'check holiday hours: as shipper, as consignee
 If rs.State <> 0 Then rs.Close
 If reqdeldte = "" Then hdte = Format$(Now, "YYYY-MM-DD") Else hdte = reqdeldte
 '** ** block shipper holiday hrs on DR for now ** SQL below looks for consignee accode ** ** ** **
     dm = "select date,site,del_op,del_cl,delappt from scops2date where accno='" & pb!cons_accnum & _
            "' and date >= '" & hdte & "' and asship <> '1' order by accno, date"
     acccnt = 0: dm3 = "CONS "

DRHOLHRS:

 rs.Open dm
 If rs.EOF Then
   If dm3 = "SHPR" Then
     dm = "select date,site,del_op,del_cl,delappt from scops2date where accno='" & pb!cons_accnum & _
          "' and date >= '" & hdte & "' and asship <> '1' order by accno, date"
     dm3 = "CONS ": GoTo DRHOLHRS
   End If
 Else
   dm2 = "": acccnt = acccnt + 1: accs(acccnt, 0) = dm3: accs2 = "": accscnt = 0
   Do
     If IsDate(rs(0)) = True Then
       If Val(rs(0)) > 0 Then
          
         dm2 = "": dm = Format$(DateValue(rs(0)), "DD-MMMYY")
         If Val(rs(1)) = 1 Then
           dm2 = "-*CLOSED*"
           If reqdeldte <> "" And dm3 = "CONS " Then 'if closed on holiday hours match with delivery date ..
             If DateValue(rs(0)) = DateValue(reqdeldte) Then opcltim = "  CLOSED"
           End If
         ElseIf Val(rs(4)) > 0 Then
           dm2 = "-APPTREQD"
         ElseIf Val(rs(2)) > 0 And Val(rs(3)) > 0 Then
           dm2 = ":" & rs(2) & "-" & rs(3)
           If reqdeldte <> "" And dm3 = "CONS " Then
             If DateValue(rs(0)) = DateValue(reqdeldte) Then opcltim = " " & rs(2) & "-" & rs(3)
           End If
         End If
         If dm2 <> "" Then
           accscnt = accscnt + 1
           If Len(accs(acccnt, 0) & dm & dm2) > 64 Then
             If Right$(accs(acccnt, 0), 2) <> ".." Then accs(acccnt, 0) = Trim$(accs(acccnt, 0)) & " .."
             accs2 = accs2 & dm & dm2 & "  "
             If accscnt > 5 Then
               accs(acccnt, 1) = "0": Exit Do
             End If
           Else
             accs(acccnt, 0) = accs(acccnt, 0) & dm & dm2 & "  "
           End If
         End If
         
       End If
     End If
     rs.MoveNext
   Loop Until rs.EOF
   If Len(accs(acccnt, 0)) > 7 Then
     accs(acccnt, 0) = Trim$(accs(acccnt, 0))
     accs(acccnt, 1) = "0" 'mark same as manual instruction
   Else
     accs(acccnt, 0) = "": accs(acccnt, 1) = "": acccnt = acccnt - 1
   End If
 End If
 rs.Close
 If dm3 = "SHPR " Then
   dm = "select date,site,del_op,del_cl,delappt from scops2date where accno='" & pb!cons_accnum & _
        "' and date >= '" & hdte & "' and asship <> '1' order by accno, date"
   dm3 = "CONS ": GoTo DRHOLHRS
 End If
 
 If rs.State <> 0 Then rs.Close
 dm = "select comment from scops2date where accno='" & pb!cons_accnum & "' and date='" & _
      Format$(Now, "YYYY-MM-DD") & "' and asship <> '1' order by accno desc limit 1"
 rs.Open dm
 If Not rs.EOF Then
   If Len(Trim$(rs(0))) > 1 Then
     acccnt = acccnt + 1: accs(acccnt, 0) = "?" & Format$(Now, "DD-MMM-YYYY") & " " & rs(0): accs(acccnt, 1) = "0"
   End If
 End If
 rs.Close
 
 '-- account action codes - consignee & shipper-linked
 msgasship = False 'default to consignee-based special instructions 11May12(DNY)
 '                0         1        2         3         4         5         6       7
 '         8         9             10            11         12        13         14
 '        15         16
 dm = "select cons_tail,cons_str,cons_cube,cons_balm,cons_roll,cons_trlr,cons_am,cons_pm," & _
      "cons_pump,cons_sideways,cons_nodblstk,cons_inside,rowascons,msgascons1,msgascons2, " & _
      "msgascons3,msgascons4 " & _
      "from scops where accno='" & pb!cons_accnum & "'"
 rs.Open dm 'consignee-based
 '                 0         1         2          3          4
 dm = "select msgasship,msgasship1,msgasship2,msgasship3,msgasship4 " & _
      "from scops where accno='" & pb!ship_accnum & "'"
 rsb.Open dm 'shipper-based messages, see below for rest of shipper codes
 If Not rsb.EOF Then
   If rsb(0) = 1 And Len(Trim$(rsb(1))) > 2 Then
     msgasship = True
     'user-entered pickup/loading instructions
     If Len(Trim$(rsb(1))) > 2 Then
       acccnt = acccnt + 1: accs(acccnt, 0) = "?" & rsb(1): accs(acccnt, 1) = "0" '0=type=manually-entered
     End If
     If Len(Trim$(rsb(2))) > 2 Then
       acccnt = acccnt + 1: accs(acccnt, 0) = "?" & rsb(2): accs(acccnt, 1) = "0"
     End If
     If Len(Trim$(rsb(3))) > 2 Then
       acccnt = acccnt + 1: accs(acccnt, 0) = "?" & rsb(3): accs(acccnt, 1) = "0"
     End If
     If Len(Trim$(rsb(4))) > 2 Then
       If acccnt < 4 Then 'first line may have been used for holiday hrs
         acccnt = acccnt + 1: accs(acccnt, 0) = "?" & rsb(4): accs(acccnt, 1) = "0"
       End If
     End If
   End If
 End If
 rsb.Close
 If Not rs.EOF Then
   If Replace(Trim$(rs(12)), "-", "") <> "" Then drrowman = rs(12)
   If Not msgasship Then
     'user-entered delivery instructions
     If Len(Trim$(rs(13))) > 2 Then
       acccnt = acccnt + 1: accs(acccnt, 0) = "?" & rs(13): accs(acccnt, 1) = "0" '0=type=manually-entered
     End If
     If Len(Trim$(rs(14))) > 2 Then
       acccnt = acccnt + 1: accs(acccnt, 0) = "?" & rs(14): accs(acccnt, 1) = "0"
     End If
     If Len(Trim$(rs(15))) > 2 Then
       acccnt = acccnt + 1: accs(acccnt, 0) = "?" & rs(15): accs(acccnt, 1) = "0"
     End If
     If Len(Trim$(rs(16))) > 2 Then
       If acccnt < 4 Then 'first line may have been used for holiday hrs
         acccnt = acccnt + 1: accs(acccnt, 0) = "?" & rs(16): accs(acccnt, 1) = "0"
       End If
     End If
   End If
   If rs(6) = 1 Then 'consignee-linked
     acccnt = acccnt + 1: accs(acccnt, 0) = "稟M DELIVERY": accs(acccnt, 1) = "1" '0=type=fixed-coded
   ElseIf rs(7) = 1 Then
     acccnt = acccnt + 1: accs(acccnt, 0) = "稰M DELIVERY": accs(acccnt, 1) = "1"
   End If
   If rs(0) = 1 Then
     acccnt = acccnt + 1: accs(acccnt, 0) = "稵AILGATE REQD": accs(acccnt, 1) = "1"
   End If
   If rs(1) = 1 Then
     If acccnt < 10 Then
       acccnt = acccnt + 1: accs(acccnt, 0) = "稴TRAIGHT TRUCK REQD": accs(acccnt, 1) = "1"
     End If
   ElseIf rs(2) = 1 Then
     If acccnt < 10 Then
       acccnt = acccnt + 1: accs(acccnt, 0) = "稢UBE VAN REQD": accs(acccnt, 1) = "1"
     End If
   ElseIf Val(rs(5)) > 10 Then
     If acccnt < 10 Then
       acccnt = acccnt + 1: accs(acccnt, 0) = "? & rs(5) & " ' TRAILER REQD": accs(acccnt, 1) = "1"
     End If
   End If
   If rs(3) = 1 Then
     If acccnt < 10 Then
       acccnt = acccnt + 1: accs(acccnt, 0) = "稨AND BALM REQD": accs(acccnt, 1) = "1"
     End If
   ElseIf rs(8) = 1 Then
     If acccnt < 10 Then
       acccnt = acccnt + 1: accs(acccnt, 0) = "稰UMP TRUCK REQD": accs(acccnt, 1) = "1"
     End If
   End If
   If rs(11) = 1 Then
     If acccnt < 10 Then
       acccnt = acccnt + 1: accs(acccnt, 0) = "稰UMP TRUCK REQD": accs(acccnt, 1) = "1"
     End If
   End If
   If rs(4) = 1 Then
     If acccnt < 10 Then
       acccnt = acccnt + 1: accs(acccnt, 0) = "稲OLL-UP DOOR REQD": accs(acccnt, 1) = "1"
     End If
   End If
   If rs(9) = 1 Then
     If acccnt < 10 Then
       acccnt = acccnt + 1: accs(acccnt, 0) = "稤O NOT LOAD SIDEWAYS": accs(acccnt, 1) = "1"
     End If
   End If
   If rs(10) = 1 Then
     If acccnt < 10 Then
       acccnt = acccnt + 1: accs(acccnt, 0) = "稤O NOT DOUBLE-STACK": accs(acccnt, 1) = "1"
     End If
   End If
 End If
 rs.Close
 If Val(pb!ship_accnum) > 0 Then 'check for account action codes linked to shipper
   dm = "select ship_haz from scops where accno='" & pb!ship_accnum & "'"
   rs.Open dm
   If Not rs.EOF Then
     If rs(0) = 1 Then 'shipper-linked
       If acccnt < 10 Then
         acccnt = acccnt + 1: accs(acccnt, 0) = "稤ANGEROUS GOODS (SHPR)": accs(acccnt, 1) = "1"
       End If
     End If
   End If
   rs.Close
 End If

SKPDRACCS:
      
 '-- final resolution of row/location - dockcheck trumps attribute trumps routing table lookup
 If drrow <> "MONTREAL" Then 'non-MTL user so show linehaul routing instead of local MTL dock info (see above)
   If western Then 'And drscac = "" Then  'temporary - drscac change to RSUT Rosensau June 1, 2020
     ''fall-thru - drrow should be WESTERN CDA
     '21Jun19 add dock check to row display ??
   ElseIf drrowdc <> "" Then
     If drrow <> "KIDY-W" Then drrow = drrowdc   '25Jun22    '"RSUT-82" Then drrow = drrowdc
   ElseIf drrowman <> "" Then
     If drrow <> "KIDY-W" Then drrow = drrowman  '25Jun22 '"RSUT-82" Then drrow = drrowman
   Else
     If drrow = "" And (ontque Or atlantic Or western) Then drrow = "SEE OPS"
   End If
 Else 'MONTREAL   21Jun19 add dock check to row display ??
 
 End If
   
   
 '2. ====== PRINT START ===========================================================================
   
 Printer.Duplex = 2 '** this has to be first printer command in a 'doc'
 Printer.Orientation = 1 'force portrait mode
 Printer.PaperSize = vbPRPSLetter  '8.5 x 11
 Printer.ScaleMode = 5 'inches
 Printer.FontTransparent = True
 
 'set vars to allow width sizing of PO list in next code segment
 Printer.FontName = "Consolas": Printer.FontSize = 10
 drpo = ""
 rs.Open "select po from pro_po where pro='" & pro & "'"  'get PO#(s) if any
 If Not rs.EOF Then
   drpo = rs(0): rs.MoveNext
   Do While Not rs.EOF
     If Printer.TextWidth(drpo & ", " & rs(0)) > 5.5625 Then '
       drpo = drpo & " ..": Exit Do
     End If
     drpo = drpo & ", " & rs(0): rs.MoveNext
   Loop
 End If
 rs.Close
 
 For pgn = 1 To pgc 'cycle thru print pages
   
   'initialize print page setup
   Printer.FontTransparent = True: Printer.DrawWidth = 3 '3 = safe thinnest line for all printers tested
   Printer.FillColor = &H0: Printer.FillStyle = 0: Printer.ForeColor = &H0 'fillstyle = solid color &H0=black
   
   '------ WATERMARK GRAPHICS ------------------------------------------------------------------
   
   'DEBUG cbsac = False: csa4 = False: drinbond = True
   
   'print 'IN BOND' horizontal in-line with 'stamp' graphic
   If drinbond Then
     Printer.FontName = "Arial Black": Printer.FontSize = 84
     Printer.ForeColor = &HE3E3E3: Printer.FontBold = True '&HE3E3E3 darkest non text-interfering grey
     '+lprnt 0.925, 2.75 + 0.0625, "IN BOND"
     lprnt 0.925, 2.75 + 0.03125 + 5.375, "IN BOND" '+
     Printer.FontName = "Arial": Printer.FontBold = False: Printer.FontSize = 11
   'if not IN BOND then check for COLLECT
   ElseIf Val(pb!Collect) > 1 Then
     Printer.FontName = "Arial Black": Printer.FontSize = 84
     Printer.ForeColor = &HE3E3E3: Printer.FontBold = True
     '+rprnt 6.5, 2.75 + 0.0625, "COLLECT"
     rprnt 6.5, 2.75 + 0.03125 + 5.375, "COLLECT" '+
     Printer.FontName = "Arial": Printer.FontBold = False: Printer.FontSize = 11
   End If
   If Not usnorthbnd And drscac = "" Then 'domestic shipment, apply spdy logo watermark
     dm = "20" 'std. Speedy watermark image blob record no. #@#
     If Val(pb!ship_accnum) = 248638 And Val(pb!billed_account) = 248638 Then dm = "25" 'Samsung watermark (with Neovia header logo) #@#
     ir.Open "select w,h,imgtype from custlogo where recno='" & dm & "'"  '#@#
     lw# = ir(0): lh# = ir(1): wlogotype = ir(2) '#@#
     ir.Close: b = GetTempLogoImage(dm, "w") '#@#
     Select Case wlogotype '#@#
       Case "wgif": barc(0).Picture = LoadPicture(sdir & "logowmimg.gif")
       Case "wjpg": barc(0).Picture = LoadPicture(sdir & "logowmimg.jpg")
     End Select
     Printer.PaintPicture barc(0), 0.85, 3.18 + 0.44, 5.73, 5.73 * lh# / lw#
     barc(0).Cls: Set barc(0).Picture = Nothing
   End If
   
   PrntProBkGrnd -1, 0, pndx, pro, typ 'pass company pre-selection (-1 = none), page y-offset, copy 1, print ctrl#
                                   'US partner Northbound bills - use table 'probusbt'
                                   'all other - determine company from Third Party Account No.
   '+PrntProBkGrnd -1, 5.375, pndx, pro, typ 'copy/page 1 or 2 (depending on partner), print ctrl#

   PrntBill pro, 0, typ, chkdig     'pro, page y-offset, orig/whse copy, check digit pro add-on (only HMES currently)
   If Not usnorthbnd Then GoTo PRINTDRSKPCSTMS
 
'4415:875 barcode pict box size for barc(0), barc2 (cdn flag) doesn't change
   
   barc(0).ScaleMode = 3 'scale graphics picture box to pixels
   If cbsac Then 'cleared
     Printer.FontName = "Arial Black": Printer.FontSize = 48: Printer.FontBold = False
     ''Printer.ForeColor = &HA0A0A0
     Printer.ForeColor = &H0
     anglePrnt "CBSA CLEARED", -0.1, 9.35, 90   '-0.03, 9.35, 90
     dm = transn & " | " & brk & " | "
     If cdte <> "" Then dm = dm & cdte
     If ctim <> "" Then dm = dm & " " & ctim & " Hrs"
     If cport <> "" Then dm = dm & " [" & cport & "]"
     If dm <> "" Then
       Printer.FontSize = 11: Printer.FontBold = True
       Printer.FontName = "Consolas": Printer.ForeColor = &H0 ''&H909090
       anglePrnt dm, 0.45, 9.2, 90
     End If
     Printer.ForeColor = &H0
   ElseIf csa4 Then 'CSA
     Printer.FontName = "Arial Black": Printer.FontSize = 48: Printer.FontBold = False
     ''Printer.ForeColor = &HA0A0A0
     Printer.ForeColor = &H0
     anglePrnt "CSA CLEARED", -0.1, 9.35, 90
     If Len(drcsano) > 8 Then
       dm = drcsano
       If cdte <> "" Then dm = dm & " | " & cdte
       If ctim <> "" Then dm = dm & " " & ctim & " Hrs"
       If cport = "" Then dm2 = entport Else dm2 = cport
       If dm2 <> "" Then dm = dm & " [" & dm2 & "]"
       Printer.FontSize = 11: Printer.FontBold = True: Printer.FontName = "Consolas"
       Printer.ForeColor = &H0 ''&H909090
       anglePrnt dm, 0.45, 9.2, 90
     End If
     Printer.ForeColor = &H0
   Else 'anything not cleared, not CSA *via RNS* record, is IN-BOND
     Printer.FontName = "Arial Black": Printer.FontSize = 48: Printer.FontBold = False: Printer.ForeColor = &H0
     anglePrnt "**", 0#, 9.35, 90
     anglePrnt "IN-BOND", -0.1, 8.65, 90
     anglePrnt "**", 0, 9.35 - 3.125, 90
     Printer.FontName = "Arial": Printer.FontSize = 14: Printer.FontBold = True
     anglePrnt indte, 0.42, 8!, 90
   End If
   barc(0).ScaleMode = 1 'twips = 1/15 pixels
   
PRINTDRSKPCSTMS:

   'pro barcode - hand-coded barcode generation/print routine
   barc(0).Cls: barc(0).BackColor = &HFFFFFF: barc(0).Width = 4415: barc(0).Height = 875
   Dim objBC As clsBarCode39    'invoke barcode print-to-picturebox class object
   Set objBC = New clsBarCode39 'call the class into memory
   With objBC 'pass properties
     .LineWeight = 1       '1-pixel base bar thickness
     .ShowLabel = False    'don't display value underneath barcode (disabled in class anyways)
     .ShowStartStop = True 'print start/top bars to improve reading
      ''If prnbtlm Or prnodfl Or prnward Then
      If prnexla Or prnward Then
        If pb!edi_unit = "" And Len(pro) = 14 Then .TextString = Right$(pro, 10) Else .TextString = pb!edi_unit
      ElseIf prnhmes Then
        If pb!edi_unit = "" Then
          If Len(pro) = 13 Then  '9Dig
            .TextString = Right$(pro, 9) & chkdig
          ElseIf Len(pro) = 14 Then
            .TextString = Right$(pro, 10) & chkdig
          Else
            .TextString = pb!edi_unit & chkdig
          End If
        End If
      ElseIf prnpyle Then
        If pb!edi_unit = "" And Len(pro) = 13 Then .TextString = Right$(pro, 9) Else .TextString = pb!edi_unit
      ElseIf prnsefl Then
        Select Case Len(pb!edi_unit)
          Case 0
            Select Case Len(pro)
              Case 11: .TextString = Mid$(pro, 5)
              Case 12
                dm = Mid$(pro, 5, 8)
                .TextString = dm & SEFLChkDig(dm)
            End Select
          Case 7, 9: .TextString = pb!edi_unit
          Case 8
            If Left$(pro, 2) = "95" Then
               .TextString = Right$(pro, 8) & SEFLChkDig(Right$(pro, 8))
            ElseIf Left$(pro, 4) = "SEFL" Then
               dm = Mid$(pro, 5, 8)
               .TextString = dm & SEFLChkDig(dm)
            End If
          Case Else: .TextString = ""
        End Select
      Else
        .TextString = pro & chkdig 'barcode OD or Speedy pro w/USF check dig if applies
      End If
     .Refresh barc(0)       'print barcode to picturebox with class code
   End With                      'x=4.125, width 5000 too long, use 3-7/8 4415 width
   Printer.PaintPicture barc(0).Image, 3.875 + 0.125, 0.84375 'print barcode in picturebox to printer (NOTE: height/width of picture box determines size/shape of output)
   Set objBC = Nothing 'clear barcode class from memory
   barc(0).Cls 'clear barcode picture box to save memory

   ''If prnodfl Or prnsefl Or prnward Or prnbtlm Then '27Aug15(DNY)-print Speedy pro barcode lower-right
   If prnhmes Or prnexla Or prnpyle Or prnward Or prnsefl Then 'print full Speedy SCAC pro lower-right
     ''x barc(0).Cls: barc(0).Width = 3650: barc(0).Height = 600 '4415, 875, object already 'Dim'ed above
     barc(0).Cls: barc(0).Width = 3650 + 900: barc(0).Height = 600 '4415, 875, object already 'Dim'ed above
     Set objBC = New clsBarCode39 'call the class into memory
     With objBC 'pass properties
       .LineWeight = 1       '1-pixel base bar thickness
       .ShowLabel = False    'don't display value underneath barcode (disabled in class anyways)
       .ShowStartStop = True 'print start/top bars to improve reading
       .TextString = pro  ''& chkdig 'barcode OD or Speedy pro w/USF check dig if applies
       .Refresh barc(0)       'print barcode to picturebox with class code
     End With                      'x=4.125, width 5000 too long, use 3-7/8 4415 width
     ''x Printer.PaintPicture barc(0).Image, 3.875 + 0.125, 8.875 'print barcode in picturebox to printer (NOTE: height/width of picture box determines size/shape of output)
     Printer.PaintPicture barc(0).Image, 3.875 + 0.125 - 0.75, 8.875 'print barcode in picturebox to printer (NOTE: height/width of picture box determines size/shape of output)
     Set objBC = Nothing 'clear barcode class from memory
     barc(0).Cls 'clear barcode picture box to save memory
   End If

PRNTDRNXTPG:
   
   If Printer.Duplex = 2 Then
     Printer.NewPage
   Else
     Printer.EndDoc: GoTo SKPPRNBOL
   End If
   
   'print BOL scan on reverse
   '** PictPlus60 .tif print routines do not work with vb duplex printing
   '   Workaround - use Windows Image Acquistion 2.0 DLL to enable picture box to display non-VB6
   '                native supported image formats such as .tif & .png, then print direct from picbox control
   '   WIA2.0 library loaded via Project ... References -> XP=system32\WIAAut.dll, Win7=syswow64\WIAAut.dll
   'If GetINI("imgparent", "0", ipardir) Then
   ipardir = "d"
   If usnorthbnd And drscac <> "RDWY" Then 'all US origin northbound except RDWY (not processed by STPC)
     dm = "select ndx from img_stpc where pro_date='" & pro & "' and inv_unit='BOL' and pgno='1' " & _
          "order by scandatetime desc limit 1"
     ir.Open dm
     If ir.EOF Then GoTo REGBOLSCAN
     dm2 = Format$(ir(0), "000000000"): ir.Close
     dm = "insert ignore into imgtemp set ndx='" & dm2 & "', img=LOAD_FILE(""/" & ipardir & "/STPCIMG/" & _
           Left$(dm2, 3) & "/" & Mid$(dm2, 4, 3) & "/" & dm2 & ".TIF"")"
     If GetPrntBOLScan(dm2, dm) Then 'error
       GoTo REGBOLSCAN
     Else
       BOLPrntDuplex
     End If
   Else 'everything else including US southbound
REGBOLSCAN:
      If ir.State <> 0 Then ir.Close
      dm = "select ndx from img_sfmaster where pro_date='" & pro & "' and ndx like 'SBOL%' and " & _
           "accnt like 'BOL-%' and pgno='1' and bkup='0' order by scandatetime desc limit 1"
      ir.Open dm
      If ir.EOF Then
         ir.Close
         dm = "select ndx from img_stpc where pro_date='" & pro & "' and inv_unit='BOL' and pgno='1' " & _
            "order by scandatetime desc limit 1"
         ir.Open dm
         If ir.EOF Then
            ir.Close
            '' supress message, per Alain Contant 28Feb2019  MsgBox "** BOL Not Found or Not Suitable to Print on DR Reverse-Side **", , "ProTrace"
            GoTo SKPPRNBOL
         End If
         dm2 = Format$(ir(0), "000000000"): ir.Close
         dm = "insert ignore into imgtemp set ndx='" & dm2 & "', img=LOAD_FILE(""/" & ipardir & "/STPCIMG/" & _
               Left$(dm2, 3) & "/" & Mid$(dm2, 4, 3) & "/" & dm2 & ".TIF"")"
         If GetPrntBOLScan(dm2, dm) Then 'error
            MsgBox "** BOL Not Found or Not Suitable to Print on DR Reverse-Side **", , "DR Print Warning"
         Else
            BOLPrntDuplex
         End If
         GoTo SKPPRNBOL
      Else
         dm2 = ir(0): ir.Close
         dm = "insert ignore into imgtemp set ndx='" & dm2 & "', img=LOAD_FILE(""/" & ipardir & "/" & _
                Left$(dm2, 8) & ".SCF/00/" & Mid$(dm2, 9, 2) & "/" & Mid$(dm2, 11, 2) & "/00" & _
                Right$(dm2, 6) & ".TIF"")"
         If GetPrntBOLScan(dm2, dm) Then 'error
            If Not usnorthbnd And drscac <> "" Then GoTo TRYDRPAPS 'US southbound shipment - try PAPS system
            MsgBox "** BOL Not Found or Not Suitable to Print on DR Reverse-Side **", , "DR Print Warning"
         Else
            BOLPrntDuplex
         End If
         GoTo SKPPRNBOL
      End If
TRYDRPAPS:
     Select Case SCAC
       Case "EXLA", "HMES", "PYLE", "RDWY", "SEFL", "WARD"
         dm = "select ndx from img_sfmaster where pro_date like '" & pro & "_' and ndx like 'PAPSMF%' and " & _
              "accnt='PAPS BOL' and pgno='1' and inv_unit='ORIG' order by scandatetime desc limit 1"
''       Case "ODFL"
''         dm = "select ndx from img_sfmaster where pro_date='" & pb!edi_unit & "' and ndx like 'PAPSMF%' and " & _
''              "accnt='PAPS BOL' and pgno='1' and inv_unit='ORIG' order by scandatetime desc limit 1"
       Case Else: GoTo SKPPRNBOL
     End Select
     ir.Open dm
     If ir.EOF Then
       ir.Close
       If Not drmnfst Then MsgBox "** BOL Scan Image Not Found to Print on DR Reverse-Side **", , "DR Print Warning"
       GoTo SKPPRNBOL
     End If
     dm2 = ir(0): ir.Close
     dm = "insert ignore into imgtemp set ndx='" & dm2 & "', img=LOAD_FILE(""/" & ipardir & "/" & _
           Left$(dm2, 8) & ".SCF/00/" & Mid$(dm2, 9, 2) & "/" & Mid$(dm2, 11, 2) & "/00" & _
           Right$(dm2, 6) & ".TIF"")"
     If GetPrntBOLScan(dm2, dm) Then
       If Not drmnfst Then MsgBox "** BOL Not Found or Not Suitable to Print on DR Reverse-Side **", , "DR Print Warning"
       GoTo SKPPRNBOL
     Else
       BOLPrntDuplex
     End If
   End If
   
SKPPRNBOL:
      
   Printer.EndDoc
   ''If pgn = pgc Then Printer.EndDoc Else Printer.NewPage 'directive to output page from printer
 Next pgn
 
 If typ = "ORIG" And drcpy < Val(gs6) Then GoTo STPDR 'control global print copy count
 
 PrntDR = CStr(pndx) 'return success with print job id#
  
PRNTDREND:
 If rs.State <> 0 Then rs.Close
 If pb.State <> 0 Then pb.Close
 If rsb.State <> 0 Then rsb.Close
 
End Function

Private Sub BOLPrntDuplex()
 Dim bolw As Integer, bolh As Integer
 Set tfi = New ImageFile 'invoke WIA 2.0 image file processing library into a newly-created object
 tfi.LoadFile sdir & "tempimg.tif": bolw = tfi.Width: bolh = tfi.Height: barc(0).Cls 'load image file, get properties
 Set barc(0).Picture = tfi.ARGBData.Picture(bolw, bolh) 'load image file into picturebox control as byte-stream
 Printer.PaintPicture barc(0).Picture, 0.2, 0.2, 7.8, 10.3 'print image from picturebox then clear picture by setting to Nothing
 Printer.FontName = "Arial Black": Printer.FontItalic = True: Printer.FontSize = 28
 Printer.FontTransparent = True: anglePrnt "***", -0.02, 8.78, 90
 anglePrnt "Consignee: Please Sign on Reverse Side", -0.09, 8.2, 90
 anglePrnt "***", -0.02, 1.97, 90: Printer.FontItalic = False: Printer.FontSize = 10
 Printer.FontName = "Arial": lprnt 0.2, 0.02, "Bill of Lading (Facsimile)"
 barc(0).Cls: Set barc(0).Picture = Nothing: Set tfi = Nothing 'end printing on current page, clean up objects & release memory
End Sub
 
Private Sub anglePrnt(rtxt$, rx!, ry!, ang&) 'rotated print independant of output device scalemode
 Dim prna As clsPrntAngle     'define a local object per angled-print class citation (see clsPrntAngle module)
 Set prna = New clsPrntAngle  'enable the object in memory
 Set prna.Device = Printer    'set the output device: Printer, form name or picturebox control name
 prna.Angle = ang: Printer.CurrentX = rx: Printer.CurrentY = ry 'set angle & left-justified' start location
 prna.PrintText rtxt 'print rotated text using the current font properties - name, size, bold, italic, etc.
 Set prna = Nothing  'remove class from memory
End Sub

Private Sub PrntBill(pro$, oy#, typ$, chkd$)
 Dim y As Double, xw As Double, xn As Double
 Dim dm As String, dm2 As String
 Dim j As Integer, k As Integer, n As Integer, k1 As Integer
 Dim b As Boolean
 Dim s() As String
  
 'white copy# embedded in top-right grey triangle - put here to get font transparency!!!
 Printer.FontName = "Arial": Printer.FontBold = True: Printer.FontSize = 20: Printer.ForeColor = &HFFFFFF
 cprnt 7.9, 0.125 + oy, "1": Printer.ForeColor = &H0
  
 Printer.FontSize = 12: Printer.FontBold = False: Printer.FontName = "Consolas"
 Printer.FontItalic = True '********* shipper/consignee in italics
   
 'consignee/shipper   12Nov13 - add contact phone#'s
 For j = 0 To 1
   y = 0.8 * (1 - j) 'reverse their locations vs. screen display
   lprnt 0.03125, 0.8 + oy + y, pb(j * 6 + 3)        'company
   lprnt 0.03125, 0.97 + oy + y, pb(j * 6 + 4)            'addr
   lprnt 0.03125, 1.14 + oy + y, pb(j * 6 + 5)        'city
   lprnt 2.25 + 0.03125, 1.14 + oy + y, pb(j * 6 + 6) 'prov
   lprnt 2.75 + 0.03125, 1.14 + oy + y, pb(j * 6 + 7) 'pcode/zip
   If Val(pb(j * 6 + 2)) > 0 Then lprnt 2.75 + 0.03125, 0.8 + oy + y, pb(j * 6 + 2) 'accnum
   dm = ""
   '''If Len(tel1(j)) > 6 Then dm = tel1(j) ' & "   "
   '''If Len(tel2(j)) > 6 Then dm = dm & tel2(j)
   '18Sep15(DNY) structured tel no. (see acodop2 - l_tel() label array, appt_tel1 scops2 field)
   If Len(tel1(j)) > 6 Then
     s = Split(tel1(j), " ")
     If UBound(s) > 1 Then
       dm = "(" & s(0) & ") " & s(1) & " " & s(2)
       If UBound(s) = 3 Then
         If Len(s(3)) > 1 Then dm = dm & " " & s(3)
       End If
     End If
   End If
   If dm <> "" Then
     Printer.FontSize = 11: lprnt 0.03125, 1.14 + 0.17 + oy + y, dm: Printer.FontSize = 12
   End If
 Next j

 Printer.FontSize = 11: Printer.FontItalic = False
 If drbyndscac <> "" Then lprnt 1.33125 + 0.47 + 0.75, 2.48 + oy, drbyndscac 'Bynd SCAC code

 Printer.FontSize = 10
 
 If usnorthbnd Then 'US northbound
   lprnt 1.932, 0.5 + oy, drscac    'orig term
   oy = 5.375
   If bt(0) <> "" Then lprnt 0.03125 + 1.5, 4.49 + oy, bt(0) 'billto/3rd party
   If Trim$(bt(1)) <> "" Then
     lprnt 0.01125, 4.625 + oy, Left$(bt(1), 35): lprnt 0.01125, 4.625 + 0.1625 + oy, bt(2)
     lprnt 0.01125, 4.625 + 0.1625 + 0.1625 + oy, bt(3)
     lprnt 0.01125 + 1.875, 4.625 + 0.1625 + 0.1625 + oy, bt(4)
     If IsNumeric(Left(bt(5), 1)) = True Then dm = Left$(bt(5), 5) Else dm = Left$(bt(5), 6)
     lprnt 0.01125 + 2.11, 4.625 + 0.1625 + 0.1625 + oy, dm
   End If
   oy = 0
   Printer.FontSize = 11
   lprnt 2.96875 + 0.47 + 0.75, 2.48 + oy, drscac 'Adv SCAC code (US partner SCAC)
   
   'advance pro ODFL -or- 'Speedy Pro' -> 31-Mar15(DNY) and WARD
   ''If prnodfl Or prnsefl Or prnward Then 'ODFL/SEFL/WARD - print their pronumber & barcode, Speedy pro prints in 'Advance Pro' area
   Select Case Left$(pro, 4)
     Case "EXLA", "HMES", "PYLE", "SEFL", "WARD"
       'speedy pro = scac + partner pro
       Printer.FontSize = 12: Printer.FontBold = True
       For j = 0 To 8
         If Printer.TextWidth(pro) < 1.25 Then Exit For
         Printer.FontSize = Printer.FontSize - 0.25
       Next j
       lprnt 3.53125 + 0.47 + 0.75, 2.48 + oy, pro
       Printer.FontSize = 11
     Case Else
       If prnsefl Or prnward Then 'SEFL/WARD
         Printer.FontSize = 12: Printer.FontBold = True
         lprnt 3.53125 + 0.47 + 0.75, 2.48 + oy, pro 'for US partner bills, this location re-named 'Speedy Pronumber'
         Printer.FontSize = 11: Printer.FontBold = False  'see below for ODFL & normal Speedy pronumber print
       Else
         If Trim$(pb!edi_unit) <> "" Then
           lprnt 3.53125 + 0.47 + 0.75, 2.48 + oy, Trim$(pb!edi_unit)
         End If
       End If
   End Select
   If Trim$(pb!agentpro) <> "" Then
     lprnt 1.83125 + 0.47 + 0.75, 2.48 + oy, Trim$(pb!agentpro)
   End If
   
   Printer.FontSize = 10
   If drtrailer <> "" Then 'northbound trailer
     If Val(drrun) > 1 Then dm = drtrailer & "-" & drrun Else dm = drtrailer
     ''If Left$(dm, 1) = "O" And drscac = "ODFL" Then dm = "OD " & Mid$(dm, 2)
     If Left$(dm, 1) = "E" And drscac = "EXLA" Then dm = "EXLA " & Mid$(dm, 2)
     If Left$(dm, 1) = "P" And drscac = "PYLE" Then dm = "PYLE " & Mid$(dm, 2)
     lprnt 0.03125, 2.48 + oy, dm 'inbound trailer
   End If
   If drdoor <> "" And UCase$(Left$(drdoor, 2)) <> "C:" Then lprnt 0.9375 + 0.47, 2.48 + oy, drdoor  'arrival door
   If entport <> "" Then lprnt 1.6875 + 0.47, 2.48 + oy, entport 'port of entry
   If ontque And drdlvyterm = "" Then 'delivery terminal not found by routing table, use inland port
     If drinport <> "" Then  'delivery terminal (should be set from Speedy routing tables when declared inland)
       Select Case Val(drinport)
          Case 395: drdlvyterm = "MT"
          Case 423: drdlvyterm = "LN"
          Case 453: drdlvyterm = "WR"
          Case 495: drdlvyterm = "TR"
       End Select
     End If
   End If
 Else 'domestic, southbound
 
   Select Case Left$(pro, 4)
     Case "EXLA", "HMES", "PYLE", "SEFL", "WARD"
       'speedy pro = scac + partner pro
       Printer.FontSize = 12: Printer.FontBold = True
       For j = 0 To 8
         If Printer.TextWidth(pro) < 1.25 Then Exit For
         Printer.FontSize = Printer.FontSize - 0.25
       Next j
       lprnt 3.53125 + 0.47 + 0.75, 2.48 + oy, pro
       Printer.FontSize = 11
     Case Else
       If prnsefl Or prnward Then 'SEFL/WARD
         Printer.FontSize = 12: Printer.FontBold = True
         lprnt 3.53125 + 0.47 + 0.75, 2.48 + oy, pro
         Printer.FontSize = 11: Printer.FontBold = False
       Else
         If Trim$(pb!edi_unit) <> "" Then
           lprnt 3.53125 + 0.47 + 0.75, 2.48 + oy, Trim$(pb!edi_unit)
         End If
       End If
   End Select
   
   oy = 5.375
   If bt(0) <> "" Then lprnt 0.03125 + 1.5, 4.49 + oy, "SZTG" & bt(0) 'billto/3rd party
   If Trim$(bt(1)) <> "" Then
     lprnt 0.03125, 4.625 + oy, Left$(bt(1), 35): lprnt 0.03125, 4.625 + 0.1625 + oy, bt(2)
     lprnt 0.03125, 4.625 + 0.1625 + 0.1625 + oy, bt(3)
     lprnt 0.03125 + 1.875, 4.625 + 0.1625 + 0.1625 + oy, bt(4)
     If IsNumeric(Left(bt(5), 1)) = True Then dm = Left$(bt(5), 5) Else dm = Left$(bt(5), 6)
     lprnt 0.03125 + 2.11, 4.625 + 0.1625 + 0.1625 + oy, dm
   End If
   oy = 0
 End If
 
 Printer.FontSize = 12: Printer.FontBold = False: Printer.FontItalic = False
 
 'service date top-right under page no.
 If reqdeldte <> "" Then
   If drbyndscac <> "NDTL" Then '08Feb12-block for NPME shipments going West via Total Logistics 'QuikX
     lprnt 6.875 + 0.15, 0.5 + oy, UCase$(Format$(DateValue(reqdeldte), "DDD-DD-MMM"))
   End If
 End If
 
 'appointment detail
 If apptstr <> "" Then
   oy = 5.375
   Printer.FontName = "Arial Black"
   Printer.FontBold = False
   Printer.FontSize = 17
   lprnt 0.03125, 4.21 + oy, UCase$(apptstr)
   Printer.FontName = "Consolas"
   Printer.FontSize = 12: Printer.FontBold = False
   oy = 0
 End If
 
 Printer.FontTransparent = True
 
 'rcvng hrs on service date
 If opcltim <> "" Then
   If InStr(opcltim, "CLOSED") > 0 Then
     lprnt 6.875 + 0.15, 0.65 + 0.0355 + oy, "**      **"
     Printer.FontBold = True: lprnt 6.875 + 0.15, 0.65 + oy, "  CLOSED": Printer.FontBold = False
   Else
     If Replace(Trim$(opcltim), "-", "") <> "" Then lprnt 6.875 + 0.12, 0.65 + oy, opcltim
   End If
 End If
 
 'rcvng hrs for week if account-coded
 If rcvnghrs Then 'box around weekly rcvng hrs
   Printer.Line (7.07, 1.5 + oy)-Step(0, -0.66): Printer.Line -Step(0.84, 0): Printer.Line -Step(0, 0.66)
   Printer.FontSize = 9
   For j = 1 To 5
     If InStr(opcl(j), "Closed") > 0 Then
       lprnt 7.12, 0.86 + ((j - 1) * 0.125) + oy, " **      **":
       lprnt 7.12, 0.84 + ((j - 1) * 0.125) + oy, opcl(j) '"M  Closed"
     Else
       lprnt 7.12, 0.84 + ((j - 1) * 0.125) + oy, opcl(j)
     End If
   Next j
 End If
 
 If acccnt > 0 Then
   Printer.FontSize = 11: Printer.FontBold = True: n = -1
   'first run-thru: manually entered instructions
   For j = 1 To acccnt
     If accs(j, 1) = "0" And Trim$(accs(j, 0)) <> "" Then
       n = n + 1: lprnt 0.84375 + 0.0625, 7.43 + (n * 0.14), accs(j, 0) 'exact next line from addinfo is 7.27375
     End If
   Next j
   If Len(accs2) > 5 Then
     n = n + 1: lprnt 0.84375 + 0.0625, 7.43 + (n * 0.14), accs2
   End If
   
   '2nd run-thru: fixed-coded
   n = -1 '': acccnt = 14
   For j = 1 To acccnt
        ''If j > 4 And accs(j, 0) = "" Then
        ''  accs(j, 0) = "? & CStr(n + 2): accs(j, 1) = "1"
        ''End If
     If accs(j, 1) = "1" And Trim$(accs(j, 0)) <> "" Then
        n = n + 1: lprnt 0.84375 + 0.0625, 8! + (n * 0.14), accs(j, 0) '7.86 exact 0.14 increments
     End If
   Next j
   
   If apptstr <> "" Then '12Jan12(DNY)
     If InStr(apptstr, "STANDING") > 0 Then dm = "稴TANDING APPOINTMENT" Else dm = "稟PPOINTMENT REQD"
     n = n + 1: lprnt 0.84375 + 0.0625, 8! + (n * 0.14), dm
   End If
   
 End If
 If loadid <> "" Then
   Printer.FontSize = 13: Printer.FontBold = True: xn = 0.105
   Do While Printer.TextWidth(loadid) > 2.25
     Printer.FontSize = Printer.FontSize - 0.5: xn = xn + 0.0075
   Loop
   lprnt 5.65625, 4.1875 + 5.375 + xn, loadid
 End If
 
 'upper-right pronumber field
 Printer.FontSize = 13: Printer.FontBold = True
 Select Case Left$(pro, 4)
   Case "EXLA"
      If edipro = "" Then edipro = Right$(pro, 10)
      lprnt 5.59375, 0.275 + oy, edipro
   Case "HMES"
      If edipro = "" Then edipro = Mid$(pro, 5)
      lprnt 5.59375, 0.275 + oy, edipro
      If chkd <> "" Then
        Printer.FontSize = 9: Printer.FontBold = False
        rprnt 6.729, 0.299 + 0.025, chkd 'add check digit if applies
        Printer.FontSize = 13: Printer.FontBold = True
      End If
   Case "PYLE"
      If edipro = "" Then edipro = Right$(pro, 9)
      lprnt 5.59375, 0.275 + oy, edipro
   Case "SEFL"
     Select Case Len(edipro)
       Case 0
         Select Case Len(pro)
           Case 11: lprnt 5.59375, 0.275 + oy, Mid$(pro, 5)
           Case 12: lprnt 5.59375, 0.275 + oy, Mid$(pro, 5) & " " & SEFLChkDig(Right$(pro, 8))
         End Select
       Case 7: lprnt 5.59375, 0.275 + oy, edipro
       Case 9: lprnt 5.59375, 0.275 + oy, Left$(edipro, 8) & " " & Right$(edipro, 1)
       Case 8: lprnt 5.59375, 0.275 + oy, edipro & " " & SEFLChkDig(Right$(pro, 8))
     End Select
   Case "WARD"
      If edipro = "" Then edipro = Right$(pro, 10)
      lprnt 5.59375, 0.275 + oy, edipro
   Case Else
      If prnward Then
        lprnt 5.59375, 0.275 + oy, edipro
      ElseIf prnsefl Then
        lprnt 5.59375, 0.275 + oy, Left$(edipro, 8) & " " & Right$(edipro, 1)
      Else
        lprnt 5.59375, 0.275 + oy, pro
        If chkd <> "" Then
          Printer.FontSize = 9: Printer.FontBold = False
          rprnt 6.729, 0.299 + 0.025, chkd 'add check digit if applies
          Printer.FontSize = 13: Printer.FontBold = True
        End If
      End If
 End Select
 Printer.FontBold = False
 
 If Trim$(pb(0)) <> "" Then 'BOL
   Do While Printer.TextWidth(pb(0)) > 1.625  '** pb!Desc - not used since 'desc' is mySql reserved word
     Printer.FontSize = Printer.FontSize - 0.5
   Loop
   lprnt 3.90625, 0.275 + oy, pb(0) 'BOL
   Printer.FontSize = 13
 End If
 If IsDate(pb!Date) = True Then
   lprnt 5.59375, 0.59 + oy, Format$(DateValue(pb!Date), "DD-MMM-YYYY") 'pro date
 End If
 If Trim$(pb!pickup) <> "" Then lprnt 3.90625, 0.59 + oy, pb!pickup   'pickup unit
 
 Printer.FontSize = 12: Printer.FontBold = False
 
 If Trim$(pb!remark) <> "" Then
   Printer.FontBold = True
   lprnt 0.84375, 2.78125 - 0.03125 + oy, pb!remark 'Remark usually = special instruction
   Printer.FontBold = False
 End If
  
 Printer.FontSize = 11: Printer.FontBold = False
 
 lprnt 4.84375 + 0.47 + 0.75, 2.48 + oy, typ & " " & reprn 'Orig/Whse, reprint count for this type
 
 'for domestic shipments only . . .
 If Not usnorthbnd Then
   Select Case pb!cons_prov
     Case "AB", "BC", "MB", "NB", "NF", "NL", "NS", "NT", "NU", "ON", "PE", "PQ", "QC", "SK", "YT"
       'origin terminal codes (pickup terminal)
       dm = ""
       Select Case pb!complete 'convert to 2-char codes
         Case "T": dm = "TR": dm = "TOR": dm2 = "(416)510-2034" 'term code, local tel.
         Case "A": dm = "AX": dm = "BRO": dm2 = "(613)525-1120"
         Case "L": dm = "LN": dm = "LON": dm2 = "(519)453-1673"
         Case "M": dm = "MT": dm = "MTL": dm2 = "(514)278-3337"
         Case "W": dm = "WR": dm = "WIN": dm2 = "(519)252-6565"
         Case "P": dm = "PI": dm = "PIC": dm2 = "(905)686-5598"
         Case "S": dm = "MI": dm = "MIS": dm2 = "(416)510-2034"
         Case "I": dm = "ML": dm = "MIL": dm2 = "(416)510-2034"
         Case "V": dm = "VA": dm = "LAC": dm2 = "(514)278-3337"
         Case Else: dm = "TOR": dm2 = "(416)510-2034"
       End Select
       lprnt 1.9375, 0.5 + oy, dm    'orig '909#
   End Select
 End If
 lprnt 3.40625, 0.5 + oy, pb!user_inits 'biller
 lprnt 2.6875, 0.5 + oy, pb!qst         'QST
 lprnt 3#, 0.5 + oy, pb!billedparty     'P/C/T/N
 lprnt 2.28125, 0.5 + oy, drdlvyterm    'DEST
 
 If Val(pb!probus) > 0 Then lprnt 3.53125, 2.5 + oy, pb!probus 'use advance pro field
 
 If oy > 0 And cpypp = 1 Then GoTo SKPPG1DRDESC
 
 'freight
 dm = "": dm2 = ""
 If Val(pb!other) > 0 And Val(pb!skids) > 0 Then
    dm = CStr(Val(pb!other) + Val(pb!skids))
    dm2 = pb!skids & " " & "PALLET(S) + " & pb!other & " " & pb!other_desc
 Else
    If Val(pb!cartons) > 0 Then
      dm = pb!cartons: dm2 = "CARTONS"
    ElseIf Val(pb!skids) > 0 Then
      dm = pb!skids: dm2 = "PALLET(S)"
    ElseIf Val(pb!drums) > 0 Then
      dm = pb!drums: dm2 = "DRUMS"
    ElseIf Val(pb!pails) > 0 Then
      dm = pb!pails: dm2 = "TRUCKLOAD"
    ElseIf Val(pb!other) > 0 Then
      dm = pb!other: dm2 = pb!other_desc
    End If
 End If
 cprnt 0.3125, 2.9375 - 0.03125 + oy, dm   'pcs count
 If Not usnorthbnd Or Left(pb!pronumber, 4) = "HMES" Or Left(pb!pronumber, 4) = "YELL" Then lprnt 0.84375, 2.9375 - 0.03125 + oy, dm2 'pcs type
 If Val(pb!asweight) > 0 Then lprnt 2.5 + 1.625, 2.78125 + oy, "CUBE-WT: " & pb!asweight
 If Val(pb!cube_feet) > 0 Then lprnt 3.75 + 1.625, 2.78125 + oy, "CUBE-FT: " & pb!cube_feet
 cprnt 6.625 + 0.53125 / 2#, 2.9375 - 0.03125 + oy, pb!totalwt
 
 'description, haz codes
 If pb!ship_accnum = "100297" And Trim$(pb!desc1) = "" Then dm = "ELECTRONICS OR RELATED" Else dm = Trim$(pb!desc1)
 If dm <> "" Then lprnt 0.84375, 3.1 - 0.0625 + oy, dm  'descrip
 
 'PO#(s) list - may be truncated - size for Consolas-10pt
 If drpo <> "" Then
   oy = 5.375: Printer.FontSize = 10: lprnt 0.84375, 4.1875 + oy - 0.1625, "PO: " & drpo: Printer.FontSize = 11: oy = 0
 End If

 j = (pgn - 1) * 30: n = 0
 For i = j To j + 29
   If i > UBound(pgi) Then Exit For
   Select Case drscac
     Case "HMES" 'remove lines for Holland bills with pricing info - per USFC Cheryl Saxton e-mail forwarded from Jared Martin 29Dec11 11:43AM
       If InStr(pgi(i), "DISCOUNT") > 0 Then 'includes requested 'HIDDEN DISCOUNT'
         GoTo NXTPGI
       ElseIf InStr(pgi(i), "ALLOWANCE") > 0 Then
         GoTo NXTPGI
       ElseIf InStr(pgi(i), "AGGREGATE SHIPMENT") > 0 Then
         GoTo NXTPGI
       ElseIf InStr(pgi(i), "FUEL SURCHARGE") > 0 Then
         GoTo NXTPGI
       ElseIf InStr(pgi(i), "NO CHARGE FOR") > 0 Then
         pgi(i) = Trim$(Replace(pgi(i), "NO CHARGE FOR", ""))
         If pgi(i) = "" Then GoTo NXTPGI
       End If
   End Select
   n = n + 1
   If pb!billed_account = "500086" Then 'ULine
     dm = Replace(pgi(i), vbLf, " ")
     If Len(dm) > 60 Then
       dm = Mid$(dm, 1, InStr(60, pgi(i), Chr(32)) - 1)
       lprnt 0.9375, 3.04 + oy + (n * 0.137), dm
       n = n + 1
       dm = Replace(pgi(i), vbLf, " ")
       dm = Mid$(dm, InStr(InStr(60, pgi(i), Chr(32)) - 1, dm, Chr(32)) + 1)
     End If
   Else
     dm = pgi(i)
   End If
   lprnt 0.9375, 3.04 + oy + (n * 0.137), dm
NXTPGI:
 Next i
  
SKPPG1DRDESC:

 If frzbl Then
   Printer.FontName = "WingDings": Printer.FontSize = 48: Printer.FontBold = True
   cprnt 3.6, 1.73, "T"  'snowflake
   Printer.FontName = "Arial": Printer.FontSize = 8: Printer.FontBold = True
   cprnt 3.6, 1.7, "FREEZABLE"
   Printer.FontName = "Consolas": Printer.FontSize = 11: Printer.FontBold = False
 End If

 If Val(pb!Collect) > 0 Then 'COLLECT
   Printer.FontName = "Arial Black": Printer.FontSize = 12: Printer.FontBold = False
   rprnt 4.75, 1.5 + oy, "COLLECT"
   lprnt 4.75 + 0.03125, 1.5 + oy, "$" & pb!Collect
   Printer.FontName = "Consolas"
 End If

 '-- row, load on storage trailer for appointment req'd consignees with future appointments
 If apptreqd Then
   If apptdate = DateValue("1900-01-01") Or apptdate > DateValue(Now) Then
     Printer.Line (6#, 1.55)-(8#, 1.8), , BF
     Printer.FillStyle = vbFSTransparent
     Printer.FontName = "ArialBlack": Printer.FontSize = 14: Printer.FontBold = True: Printer.FontItalic = True
     Printer.ForeColor = &HFFFFFF
     cprnt 7#, 1.57, "LOAD ON STORAGE"
     Printer.ForeColor = &H0: Printer.FontItalic = False
   End If
 End If
 Printer.FontName = "Arial": Printer.FontBold = True: Printer.FontSize = 36
 rprnt 7.95, 1.8 + oy, drrow
 Printer.FontSize = 13: Printer.FontName = "Consolas"
 
 b = False
 Select Case drrowdc '12Jan21(DNY)
   Case "EUGHOLD", "JOTIHOLD", "KENHOLD", "R4ANDRE": b = True
 End Select
 ''If drrowdc = "R4ANDRE" Or drrowdc = "JOTIHOLD" Or drrow = "RSUT-82" Then '28May20
 If b Or drrow = "KIDY-W" Then  '25Jun22  '"RSUT-82" Then '12Jan21(DNY)
   '24Oct19(DNY) already printed above as 'drrow' set in sub PrntDR
   '26Feb20(DNY) added GOJO routing change from CFF to West 82
 Else
   If drdlvyterm <> "TR" And drrowdc <> "" Then
     rprnt 7.95, 1.55 + oy, "[" & drrowdct & ":" & drrowdc & "]"
   End If
 End If

 Printer.FontSize = 10: Printer.FontBold = False

End Sub

Private Sub PrntProBkGrnd(cpy%, oy#, pdx&, pro$, typ$)
 Dim dm As String
 Dim b As Boolean
 Dim x1arg As Long, y1arg As Long
 Dim x2arg As Long, y2arg As Long
 Dim x3arg As Long, y3arg As Long
 
 Printer.FontTransparent = True
 Printer.ForeColor = &H0: Printer.FillColor = &H0: Printer.FillStyle = 0
 Printer.FontName = "Arial": Printer.FontBold = True
 Printer.FontItalic = False: Printer.FontUnderline = False
   
 'top-right corner .575w x .36h" grey triangle
 If oy <= 0 Then cpypp = 2 'default 2 copies per page (upper & lower) - form-level var
 x1arg = Printer.ScaleX(7.42, vbInches, vbPixels): y1arg = Printer.ScaleY(0.145, vbInches, vbPixels)
 x2arg = Printer.ScaleX(8.005, vbInches, vbPixels): y3arg = Printer.ScaleY(0.505, vbInches, vbPixels)
 FillTriangle Printer.hdc, x1arg, y1arg, x2arg, y1arg, x2arg, y3arg, &H808080 'API-based, global.bas
 ''Set pic2.Picture = LoadPicture(drsdir & "trcrnrbk.JPG"): Printer.PaintPicture pic2, 7.425, 0.14 + oy, 0.575, 0.36
 ''Set pic2.Picture = Nothing
   
 '-- LOGO
 If logorec = "1" And dr_tel <> "" Then
   Set pic2.Picture = LoadPicture(drsdir & "speedy.jpg")
   Printer.PaintPicture pic2, 0!, 0.015 + oy, 1.8, 0.554 'speedytransport w x h aspect: 3.25
 ElseIf logorec > 1 And dr_tel <> "" Then
   Select Case logotype '#@#
     Case "gif": Set pic2.Picture = LoadPicture(sdir & "logoimg.gif")
     Case "jpg": Set pic2.Picture = LoadPicture(sdir & "logoimg.jpg")
   End Select
   Printer.PaintPicture pic2, logox, logoy + oy, logow, logoh
 End If
 
    oy = 0 '+\\_//^\\_//
 
 'page # of Tot
 Printer.FontName = "Arial": Printer.FontBold = True
 Printer.FontSize = 15
 rprnt 7# + 0.03125, 0.1875 - 0.03125, CStr(pgn)
 lprnt 7# + 0.25, 0.1875 - 0.03125, CStr(pgc)
 rprnt 7# + 0.03125, 0.1875 - 0.03125 + oy, CStr(pgn)
 lprnt 7# + 0.25, 0.1875 - 0.03125 + oy, CStr(pgc)
 Printer.FontBold = False: Printer.FontSize = 11
 cprnt 7# + 0.125 + 0.03125 / 2!, 0.1875 + 0.03125 / 2!, "of"
 cprnt 7# + 0.125 + 0.03125 / 2!, 0.1875 + (0.03125 / 2!) + oy, "of"
 
 'title text
 Printer.FontName = "Arial": Printer.FontSize = 7 'font=Arial 7pt Bold
 Printer.FontBold = True
 'main address
 Select Case cpy
   Case -1 'Speedy
     Select Case logorec
       Case 1: lprnt 3.875, 0.03125 + oy, "(800) 265-5351" 'w/Speedy old logo
       Case 13: lprnt 3.875, 0.03125 + oy, "(800) 265-5351" 'new STG logo
       Case 50  'EXLA
         If dr_tel <> "" Then lprnt 3.875, 0.03125 + oy, dr_tel
         lprnt 1.90625, 0.09375 + oy, "P.O. BOX 25612, Richmond, VA 23260"
         lprnt 1.90625, 0.1975 + oy, "www.estes-express.com"
         GoTo SKPPRADDR
       Case Else 'partner billto's
         If dr_tel <> "" Then lprnt 3.875, 0.03125 + oy, dr_tel
     End Select
     lprnt 1.90625, 0.09375 + oy, "265 RUTHERFORD ROAD SOUTH"
     lprnt 1.90625, 0.1975 + oy, "BRAMPTON, ONTARIO   L6W 1V9"
 End Select
 
SKPPRADDR:
     
 Printer.FontBold = False
 rprnt 7.95, 0.03125 + oy, "DELIVERY RECEIPT"
 
 '=+=+=+=+  program-specific +=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+
   If init = "" Then dm = "   " Else dm = init
 '=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+=+
 
 dm = Format$(pdx, "0000#") & dm & UCase$(Format$(Now, "DDMMMYY-HhNn"))
 lprnt 3.875 + 2.875 - Printer.TextWidth(dm), 0.03125 + oy, dm
     
 Printer.FontSize = 6
 Printer.FontBold = False
 lprnt 0.03125, 0.6875 + oy, "CONSIGNEE"
 lprnt 0.03125, 1.5 + oy, "SHIPPER"
 
 lprnt 1.9375, 0.375 + oy, "ORIG."
 lprnt 2.28125, 0.375 + oy, "DEST."
 lprnt 2.6875, 0.375 + oy, "QST"
 lprnt 3#, 0.375 + oy, "P/C/T/N"
 lprnt 3.40625, 0.375 + oy, "BILLER"
 
 lprnt 3.90625, 0.1875 + oy, "SHIPPER BOL NO."
 lprnt 3.90625, 0.5 + oy, "PICKUP UNIT"
 
 dm = "PRO NUMBER"
 If prnexla Then dm = "EXLA " & dm
 If prnhmes Then dm = "HMES " & dm
 If prnpyle Then dm = "PYLE " & dm
 If prnsefl Then dm = "SEFL " & dm
 If prnward Then dm = "WARD " & dm
' lprnt 5.59375, 0.1875 + oy, dm
 
 lprnt 5.59375, 0.5 + oy, "DATE"
 
 lprnt 0, 2.375 + oy, "INBOUND TRAILER"
 lprnt 0.9375 + 0.47, 2.375 + oy, "DOOR"
 lprnt 1.6875 + 0.47, 2.375 + oy, "PORT"
 
 lprnt 1.33125 + 0.47 + 0.75, 2.375 + oy, "BYD.SCAC"
 lprnt 1.83125 + 0.47 + 0.75, 2.375 + oy, "BEYOND PRO"
 lprnt 2.96875 + 0.47 + 0.75, 2.375 + oy, "ADV.SCAC"
 
 ''If prnodfl Or prnsefl Or prnward Or prnbtlm Then lprnt 3.53125 + 0.47 + 0.75, 2.375 + oy, "SPEEDY PRONUMBER" Else lprnt 3.53125 + 0.47 + 0.75, 2.375 + oy, "ADVANCE PRO"
 If prnhmes Or prnexla Or prnsefl Or prnward Then lprnt 3.53125 + 0.47 + 0.75, 2.375 + oy, "SZTG PRONUMBER" Else lprnt 3.53125 + 0.47 + 0.75, 2.375 + oy, "ADVANCE PRO"
 
 Printer.FontSize = 8
 Select Case Left$(pro, 4)
   Case "EXLA", "HMES", "PYLE", "SEFL", "WARD"
     cprnt 5.175 - 0.3725, 8.725, pro
   Case Else
     If prnsefl Or prnward Then
       cprnt 5.175 - 0.3725, 8.725, "SZTG " & pro
     End If
 End Select
 Printer.FontSize = 6
 lprnt 4.84375 + 0.47 + 0.75, 2.375 + oy, "PRINT RECORD"

 cprnt 0.3125, 2.6875 + oy, "PIECES"
 cprnt 0.625 + (0.8125 - 0.625) / 2#, 2.6875 + oy, "H/M"
 lprnt 0.8125 + 0.0625, 2.6875 + oy, "DESCRIPTION"
 cprnt 6.625 + 0.5625 / 2#, 2.6875 + oy, "WEIGHT(LBS)"
 lprnt 6.625 + 0.5625 + 0.25, 2.6875 + oy, "CHARGES"
 
    oy = 5.375 '+\\_//^\\_//
  
 lprnt 0.0318, 4.1875 + oy, "APPT. DATE"
 lprnt 1.05 + 0.465 + 0.365, 4.1875 + oy, "APPT. TIME(S)"
 lprnt 2.075 + 0.525 + 0.66, 4.1875 + oy, "NO."
   
 lprnt 5.65625, 4.1875 + oy, "LOAD-ID / CONFIRMATION NO." 'chngd 05Feb14 per JAM
  ''lprnt 5.65625, 4.1875 + oy, "CPC CONTROL NO.": lprnt 6.65625, 4.1875 + oy, "OUT"
  ''lprnt 7.109375, 4.1875 + oy, "RETURN": lprnt 7.5625, 4.1875 + oy, "CPC CLSD"

 lprnt 0.03125, 4.5 + oy, "THIRD PARTY"
 lprnt 2.78125, 4.640625 + oy, "FIRM"
'&& lprnt 2.78125, 4.859375 + oy, "PRINT NAME"
 lprnt 2.78125, 4.859375 + oy, "NAME"
'&& lprnt 2.78125, 5.078125 + oy, "SIGNATURE: RECORDED BY STYLUS ON MOBILE DEVICE - SEE ABOVE"

 lprnt 5.65625, 4.5 + oy, "IN"
 lprnt 5.65625, 4.8125 + oy, "OUT"
 lprnt 6.25, 4.734375 + oy, "DRIVER"
 lprnt 6.25, 4.984375 + oy, "DATE DEL'D"
   
 'different font texts
 Printer.FontSize = 7
 Printer.FontTransparent = False: Printer.ForeColor = &H606060
 lprnt 0.8125 + 0.4, 7.32, "  DELIVERY INSTRUCTIONS  "
 Printer.FontTransparent = True: Printer.ForeColor = &H0
 
 'yellow-sticker replacement for Warehouse copies only
 If typ = "WHSE" Then
   Printer.FontName = "Arial Black": Printer.FontSize = 128: cprnt 3.75, 5.2, Right$(pro, 4)
 End If
 
 Printer.FontSize = 5.55
 Printer.FontName = "Arial": Printer.FontBold = False
 Printer.FontSize = 5
 cprnt 7.531253, 4.8125 + oy, "DD" 'date del'd format
 lprnt 7.15625, 4.8125 + oy, "MM"
 lprnt 7.78125, 4.8125 + oy, "YY"
 lprnt 7.45, 5.078125 + oy, "REV.220719DNY"
 
 '-- Lines -----------------------------------------------------------------------

 oy = 0
 
 Printer.Line (0, 0 + oy)-Step(0, 0.0625), &H0 'mark bill top corners, left
 Printer.Line (0, 0 + oy)-Step(0.0625, 0)
 Printer.Line (8#, 0 + oy)-Step(0, 0.0625) 'right
 Printer.Line (8#, 0 + oy)-Step(-0.0625, 0)
 
 Printer.Line (3.875, 0.1875 + oy)-Step(2.875, 0) 'bol, pro top
 Printer.Line (3.875, 0.5 + oy)-Step(2.875, 0)    'po# top
 Printer.Line (3.875, 0.8125 + oy)-Step(2.875, 0) 'bottom
 
 Printer.Line (1.90625, 0.375 + oy)-Step(1.96875, 0)  'orig, dest top
 Printer.Line (0, 0.6875 + oy)-Step(3.875, 0) 'consignee top
 Printer.Line (0, 1.5 + oy)-Step(8#, 0)  'shipper top
 
 Printer.Line (0, 2.375 + oy)-Step(8#, 0)     'top 2 lines of freight area
 Printer.Line (0, 2.6875 + oy)-Step(8#, 0)
 
    oy = 5.375 '+\\_//^\\_//
 
 Printer.Line (0, 4.1875 + oy)-Step(8#, 0)    'bottom freight area
 Printer.Line (0, 4.5 + oy)-Step(8#, 0)
 Printer.Line (0, 5.15625 + oy)-Step(2.75, 0) 'left bottom
 
 Printer.Line (5.625, 4.8125 + oy)-Step(0.59375, 0) 'in (bottom)
 Printer.Line (5.625, 5.0625 + oy)-Step(0.59375, 0) 'out (bottom)
  
 Printer.Line (6.59375, 4.8125 + oy)-Step(1.4375, 0) 'driver
 Printer.Line (6.75, 5.0625 + oy)-Step(1.25, 0)      'date delvd
 
 Printer.Line (3.015625, 4.71875 + oy)-Step(2.609375, 0) 'firm
 '&& Printer.Line (3.3125, 4.9375 + oy)-Step(2.3125, 0)      'printname
 Printer.Line (3.3125 - 0.27, 4.9375 + oy)-Step(2.3125 + 0.27, 0)
 'Printer.Line (3.28125, 5.15625 + oy)-Step(2.34375, 0)   'signature
 
    oy = 0 '+\\_//^\\_//
 
 Printer.Line (0.8125 + 1.75, 7.38)-(6.625, 7.38), &H808080 'delivery instructions area separator
 Printer.Line (0.8125, 7.38)-(0.8125 + 0.41, 7.38), &H808080
 
 'all verticals from left ------------------------
 Printer.Line (0, 0.6875 + oy)-Step(0, 1.25)      'consignee
  
 Printer.Line (3.875, 0.1875 + oy)-Step(0, 0.625)  'bol, pkup unit left
 Printer.Line (5.5625, 0.1875 + oy)-Step(0, 0.625) 'mid line
 Printer.Line (6.75, 0.1875 + oy)-Step(0, 0.625)   'pro,date right
 
 Printer.Line (1.90625, 0.375 + oy)-Step(0, 0.3125) 'orig
 Printer.Line (2.25, 0.375 + oy)-Step(0, 0.3125)    'dest
 Printer.Line (2.65625, 0.375 + oy)-Step(0, 0.3125) 'QST
 Printer.Line (2.96875, 0.375 + oy)-Step(0, 0.3125) 'P/C/T/N
 Printer.Line (3.375, 0.375 + oy)-Step(0, 0.3125)   'biller
 
 Printer.Line (0.90625 + 0.47, 2.375 + oy)-Step(0, 0.3125)       'door
 Printer.Line (1.65625 + 0.47, 2.375 + oy)-Step(0, 0.3125)       'port
 
 Printer.Line (1.3 + 0.47 + 0.75, 2.375 + oy)-Step(0, 0.3125)    'byd scac
 Printer.Line (1.8 + 0.47 + 0.75, 2.375 + oy)-Step(0, 0.3125)    'new: Bynd Pro  'beyond rev
 Printer.Line (2.9375 + 0.47 + 0.75, 2.375 + oy)-Step(0, 0.3125) 'adv scac
 Printer.Line (3.5 + 0.47 + 0.75, 2.375 + oy)-Step(0, 0.3125)    'adv pro
 Printer.Line (4.8125 + 0.47 + 0.75, 2.375 + oy)-Step(0, 0.3125) 'adv date
 
    oy = 5.375 '+\\_//^\\_//
 
 Printer.Line (2.75, 4.5 + oy)-(2.75, 5.15625 + oy) 'firm, printname signature
  
 Printer.Line (0.625, 2.6875)-(0.625, 4.1875 + oy) 'pcs
 Printer.Line (0.8125, 2.6875)-(0.8125, 4.1875 + oy)  'H/M
 Printer.Line (6.625, 2.6875)-(6.625, 4.1875 + oy)       'was chgs, now tot wt
 Printer.Line (6.625 + 0.5625, 2.6875)-(6.625 + 0.5625, 4.1875 + oy)  'new charges
  
 Printer.Line (5.625, 4.1875 + oy)-Step(0, 0.875)     'right of Recv'd in good . . .
 Printer.Line (6.21875, 4.1875 + 0.3125 + oy)-Step(0, 0.96875 - 0.3125) 'left of driver, date del'd
 
 Printer.Line (7.3125, 5.0625 + oy)-Step(0.125, -0.21875) 'date del'd separators
 Printer.Line (7.625, 5.0625 + oy)-Step(0.125, -0.21875)
 
End Sub

Private Function GetPrntBOLScan(ndx$, dm$) As Boolean
 Dim icmd As New ADODB.Command
 Dim ist As New ADODB.Stream
 Dim isc As New ADODB.Stream
 On Error Resume Next
 Err = 0: ist.Open: isc.Open  'init ADO stream objects as ubiquitous streams to start
 isc.Type = adTypeBinary      'set stream to binary (auto-converts text to bin data during 'copyto' op)
 icmd.ActiveConnection = ir.ActiveConnection: icmd.CommandType = adCmdText
 On Error GoTo GPBSErr
 icmd.CommandText = dm: icmd.Execute ': Pause 0.1
 ir.Open "select img from imgtemp where ndx='" & ndx & "'"
 ist.Flush: ist.WriteText ir(0): ir.Close: ofs = ist.Size
 ist.SetEOS: ist.Position = 2: isc.Flush: ist.CopyTo isc, ofs - 2
 ist.Flush: isc.SaveToFile sdir & "tempimg.tif", adSaveCreateOverWrite
 ist.Close: isc.Close: GoTo GPBSEND
GPBSErr:
 On Error Resume Next
 If ir.State <> 0 Then ir.Close
 isc.Flush: GetPrntBOLScan = True
GPBSEND:
 On Error Resume Next
 icmd.CommandText = "delete from imgtemp where ndx='" & ndx & "'": icmd.Execute
 Set ist = Nothing: Set isc = Nothing: Set icmd = Nothing: Err = 0
End Function


Private Sub SendMail(subj$, Msg$, mtfn$) 'send e-mail via SMTP - vars: subject line, primary message, mySQL mail-to fieldname, pronumber
 Dim j As Integer
 Dim dm As String, logmsg As String, mailto As String
 Dim mt(1 To 40) As String
 
 On Error GoTo ERSNDML
 logmsg = "" ': Set ef = fs.OpenTextFile(sdir & "ManifestMailer.log", 8, True) 'open mail-send log file

 'build mailto list
 rsc.Open "select email from shpcntrptmailst where " & mtfn & "='1' order by email"
 If rsc.EOF Then
   logmsg = Format$(Now, "DDMMMYY HhNn") & " No Mailto Recipients for Fieldname " & mtfn: rsc.Close: GoTo EXSNDML
 End If
 j = 0: Erase mt: mailto = ""
 Do
   j = j + 1: mt(j) = rsc(0): mailto = mailto & rsc(0) & "; ": rsc.MoveNext 'build individual + combined list
 Loop Until rsc.EOF Or j = 40
 rsc.Close
 If Len(mailto) > 3 Then mailto = Left$(mailto, Len(mailto) - 2)
 'subject & message passed as parameters
 
 '** 1 - connect to e-Mail server via SMTP port 25
 If wsck.State <> 0 Then
   wsck.Close: Pause 0.1
 End If
 wsck.RemotePort = 25: wsck.RemoteHost = "132.147.119.92"  'Speedy POP3 mail server
 emconn = False: wsck.Connect: logmsg = Waitemresp(0)
 If logmsg <> "" Then
   logmsg = "[Connect] " & logmsg: GoTo EXSNDML
 End If
  
 '** 2 - send test message - server responds into var 'emdata' and sets 'emevent' flag True
 emevent = False: wsck.SendData "HELO speedy.ca" & vbCrLf: logmsg = Waitemresp(1)
 If logmsg <> "" Then
   logmsg = "[HELO] " & logmsg: GoTo EXSNDML
 End If
  
 '** 3 - set e-mail sender
 emevent = False: wsck.SendData "MAIL FROM: " & "david.young@speedy.ca" & vbCrLf: logmsg = Waitemresp(1)
 If logmsg <> "" Then
   logmsg = "[From] " & logmsg: GoTo EXSNDML
 End If
  
 '** 4 - set e-mail recipient(s)
 For j = 1 To 40 'cycle thru mailto's array
   If mt(j) = "" Then Exit For
   If InStr(mt(j), "@") > 0 Then
     emevent = False: wsck.SendData "RCPT TO: " & "<" & mt(j) & ">" & vbCrLf: logmsg = Waitemresp(1)
     If logmsg <> "" Then
       rs.Close: logmsg = "[RCPT]" & mt(j) & " " & logmsg: GoTo EXSNDML
     End If
   End If
 Next j
  
 '** 5 - put server in message receive mode
 emevent = False: wsck.SendData "DATA" & vbCrLf: logmsg = Waitemresp(1)
 If logmsg <> "" Then
   logmsg = "[DATA] " & logmsg: GoTo EXSNDML
 End If
  
 '** 6 - send e-mail 'To:' list
 wsck.SendData "To: " & mailto & vbCrLf
  
 '** 7 - send subject line
 wsck.SendData "Subject: " & subj & vbCrLf & vbCrLf
  
 '** 8 - send message body
 wsck.SendData Msg & vbCrLf & vbCrLf
  
 '** 9 - end message with single '.'
 emevent = False: wsck.SendData "." & vbCrLf: logmsg = Waitemresp(1)
 If logmsg <> "" Then
   logmsg = "[.] " & logmsg: GoTo EXSNDML
 End If
  
 '** 10 - disconnect from server
 wsck.SendData "QUIT" & vbCrLf: logmsg = Format$(Now, "DDMMMYY HhNn") & " Pro " & mpro
 GoTo EXSNDML
  
ERSNDML:
 logmsg = Format$(Now, "DDMMMYY HhNn") & " Pro " & mpro & " Internal Err " & Err.Description
  
EXSNDML:
 If wsck.State <> 0 Then wsck.Close
 If logmsg <> "" Then
   'ef.WriteLine Format$(Now, "DDMMMYY HhNn") & " " & logmsg
 End If
 'ef.Close
 On Error GoTo 0
 
End Sub

Private Function Waitemresp$(parm%)
 Dim j As Integer, k As Integer
 j = 0: k = 0
 Select Case parm
   Case 0 'server connect 4 sec. response time (.05 sec intervals), max. 3 connection attempts to fail
     DoEvents
     Do Until emconn = True
       Pause 0.05: DoEvents: j = j + 1
       If j = 50 Then   '200 Then 'up to 4 sec. (200 x 0.05)
         k = k + 1
         If k > 2 Then
           Waitemresp = "FAIL Connect to e-Mail Server!": Exit Function
         End If
         j = 0: wsck.Close: DoEvents: wsck.Connect
       End If
     Loop
   Case 1
     DoEvents: j = 0
     Do Until emevent = True
       Pause 0.05: DoEvents: j = j + 1
       If j = 100 Then '500 Then
         Waitemresp = "FAIL Retreive Response from e-Mail Server!": Exit Function
       End If
     Loop
 End Select
End Function

Private Sub wsck_Close()
 emconn = False 'indicate mail server connection closed
End Sub

Private Sub wsck_Connect() 'fires on intial TCP connection with e-mail server
 emconn = True
End Sub

Private Sub wsck_DataArrival(ByVal bytesTotal&) 'e-mail server replies with data
 Dim dm As String
 On Error Resume Next
 wsck.GetData dm, vbString, bytesTotal: emdata = dm
 If emdata <> "" Then emevent = True
EXDA:
End Sub

Private Sub wsck_Error(ByVal Number%, Description$, ByVal Scode&, ByVal Source$, ByVal HelpFile$, ByVal HelpContext&, CancelDisplay As Boolean)
 'ef.WriteLine Format$(Now, "DDMMMYY HhNn") & " TCP Err " & Description
 MsgBox "TCP Err " & Description
End Sub

Public Sub Pause(sec!)
 sec = Timer + sec  'Timer in context of time elapsed since midnight
 Do
   'NOTE: no DoEvents - operation stops all other thread execution
 Loop While Timer < sec
End Sub

Private Sub trstat_Click()
 gs1 = cscac
 gs2 = Trim$(Mid$(clstc.Text, 3, 11))
 tstatfrm.Show 1
End Sub
