

require 'schmiedebasis'
PROGRAMM_MIT_DBF_ZUGRIFF = true
require 'models/system_fkt/dbase_zugriff'
require 'models/system_fkt/excel_zugriff'
require 'models/generali'
require 'models/wandler/excel_zuordnung'
require 'models/wandler/wandler_zu_excel'
require 'models/wandler/wandler_hr5'
require 'models/system_fkt/speicherung'


if __FILE__ == $0 then
  durchlaufe_unittests($0)
else


if not defined? PROZENT_Proc then

PROZENT_Proc  = proc { |wert| (wert||0) / 100 }
PROMILLE_Proc = proc { |wert| (wert||0) / 1000 }

trc_hinweis :PROGRAMM_MIT_DBF_ZUGRIFF,  PROGRAMM_MIT_DBF_ZUGRIFF

class Dbf2ExlError < RuntimeError
end

class FatalDbf2ExlError < RuntimeError
end

class Dbf2ExlErrorMappeStehtImWeg < Dbf2ExlError
end

$wandel_art = :NG1

def dbf2exl_allgemein_stop()
  $exlist_wandler = nil
  $wandel_art = nil
  "FERTIG"
end

def dbf2exl_allgemein_start(waart_str, ziel_pfad)
  $wandel_art = waart_str.to_sym
end


def dbf2exl_hr6_start(ziel_pfad)
  $wandel_art = :HR6
  $exlist_info_geschrieben = false
  begin
	  mappe = ExcelZugriff.excel_vorlage_oeffnen($wandel_art)
	  $exlist_wandler = WandlerZuExcel.new(ziel_pfad, mappe)
    
    $storno_prozeduren ||= []
    $storno_prozeduren << proc {dbf2exl_hr6_stop}
  rescue
    trc_aktuellen_error "hr6_start"
    dbf2exl_hr6_stop rescue nil
    raise FatalDbf2ExlError, "Init-Fehler: #{$!}"
  end
end

def dbf2exl_hr6_stop()
  #schreibe_info(info_hash)
  $exlist_wandler.finalisiere if $exlist_wandler
  $exlist_wandler = nil
  $wandel_art = nil
  "FERTIG"
end

# vsnr_pfad ist /pfad_zu_profil/profil/vsnr
def dbf2exl_vsnr(vsnr_pfad, ziel_pfad, mappen_offenhalten=false)
  trc_info "wandel_art, vsnr_pfad", [$wandel_art, vsnr_pfad]
  begin
    #vsnr_pfad = vsnr_pfad.gsub(/[<>]/,"")
    ziel_pfad = ziel_pfad.gsub(/[<>]/,"")
    trc_info :dbf2exlvsnr_ziel, ziel_pfad

    DbfDat::oeffnen(File.dirname(vsnr_pfad), :s)

    vsnr = File.basename(vsnr_pfad)
    trc_temp :erz_excel__vsnr=, vsnr


    vd = DbfDat::Vd.find( :first, "vsnr" => vsnr)
    if !vd then
      trc_fehler "vd nicht gefunden", vsnr_pfad
      raise "dbf-Vertrag (#{vsnr_pfad}) nicht gefunden!"
    end
    trc_info :vd_found, vd.vsnr

    startzeit = Time.now
    ExcelZugriff.sichtbar = KONFIG.opts[:MappenAnzeigenBeiBatchjobs]
    
    case $wandel_art.to_s 
    when "HR6" then         
      $exlist_wandler.starte_wandlung([vd])
      unless $exlist_info_geschrieben then
        schreibe_info(vsnr_pfad)
        $exlist_info_geschrieben = true
      end
    when "HR5"
      wandler = Wandler_Hr5.new(File.dirname(vsnr_pfad), File.dirname(ziel_pfad), :mappen_offenhalten => mappen_offenhalten)
      wandler.wandle_vertrag(vsnr, File.basename(ziel_pfad))
    when "NG1" then
	    mappe = dbf2exl_vorgaben(vd, ziel_pfad, vsnr_pfad)
	    mappe.Close if mappe and not mappen_offenhalten
	
	    ref_dateiname = ziel_pfad.dup
	    if ref_dateiname.sub!(/\/([^\/]+)$/,'/Referenzdaten/'+ExcelZugriff::STD_REFPREFIX+'\1')
	      Phi::force_dirs(File.dirname(ref_dateiname))
	      trc_info :ref_dateiname, ref_dateiname
	
	      refmappe = dbase_in_vorlagenmappe(vd, ref_dateiname)
	      refmappe.Close if refmappe and not mappen_offenhalten
	    end
	  else
      raiese "Bug: unbekannte Wandlungs-Art: #{$wandel_art.inspect}"
	  end
    
    stopzeit = Time.now
    dauer_gespeichert = stopzeit - startzeit
  ensure
    DbfDat::schliessen
  end
  dauer_gespeichert
end

def schreibe_info(vsnr_pfad)
  if vsnr_pfad then
	  vsnr = File.basename(vsnr_pfad)
	  kurzinfo = begin
	    ort_persist = OrtPersistenz.new(File.dirname(vsnr_pfad))
	    ort_persist.inhalt_fuer_eintrag(vsnr).kurzinfo
	  rescue
	    trc_aktuellen_error :kurzinfo_lesen
	    "-"
	  end
	
	  ordnerinfo = begin
	    ort_persist.inhalt_fuer_eintrag(".").kurzinfo
	  rescue
	    trc_aktuellen_error :ordnerinfo_lesen
	    "."
	  end
	
	  info_hash = {
	    "KurzInfo" => kurzinfo,
	    "GruppenInfo" => ordnerinfo,
	    "DateiQuelle" => vsnr_pfad
	  } 
  else
    info_hash = { } 
  end       
  info_hash["DateiHistorie"] = 
    Time.now.strftime("%y-%m-%d: ") + "Erstellt mit ATS #{Ats::VERSION::STRING}."   
       
  info_hash.each do |bezug, wert|
    ExcelZugriff.sicheres_schreiben(bezug, wert) if wert
  end
  
end


def dbf2exl_vorgaben(vd, dateiname, vsnr_pfad=nil)
  trc_hinweis :dbf2exl_vorgaben_start   
  
  begin
    case $wandel_art.to_s
    when "NG1"
      mappe = ExcelZugriff.excel_vorlage_oeffnen("NG1")
      trc_info :mappe_erst_start
	    ExcelZugriff.speichere_kontrolliert(mappe,
	                                        dateiname,
	                                        KONFIG.opts[:WennBeiExcelErstellungMappeExistiert])
	    vd.vorgaben_rekursiv_eintragen(mappe)
	    trc_info :NG1_fertig, dateiname
      schreibe_info(vsnr_pfad) 
    else
      raise "Nur NG1 erlaubt, aber: #{$wandel_art}"  
#    when "HR6"
 #     $exlist_wa.starte_wandlung([vd])
  #    trc_info :HR6_fertig_vsnr=, vd.vsnr
    end
  rescue
    trc_aktuellen_error "Problem beim Erstellen der Vorgaben-Arbeitsmappe"
    raise "Problem beim Erstellen der Vorgaben-Arbeitsmappe (\"#{
                                   $!.to_s.split("\n")[0][0..60].strip}\")"
  end
  
  mappe.save
  mappe

  #File.open("#{WORK_DIRNAME}/tmp.yaml","w") {|f| f.puts vd.to_yaml}

end

$kurze_liste = true
if $kurze_liste then
  $erste_zeile = 6
else
  $erste_zeile = 2
  raise "Bug: Alte Version der Vorlagenmappen nicht mehr unterstï¿½ï¿½zt"
end

def dbase_in_vorlagenmappe(vd, dateiname)
      trc_hinweis :dbase_invorlmappe_start
      xl = ExcelZugriff.application
      begin
        vorl_dateiname = 'E2RefVorl.xls' #KONFIG.opts[:ExstarVorlage_DateiName]
        vorl_dateiname = ORIG_DIRNAME + "/" + vorl_dateiname if vorl_dateiname !~ /\//
        trc_info :vorl_datnam, vorl_dateiname
        mappe = xl.Workbooks.Open(vorl_dateiname.gsub("/","\\"))
        trc_info :vorl_mappe, mappe
        aktion_falls_schon_existiert = KONFIG.opts[:WennBeiExcelErstellungMappeExistiert]
        ExcelZugriff.speichere_kontrolliert(mappe, dateiname, aktion_falls_schon_existiert)
      rescue WIN32OLERuntimeError
        trc_hinweis "Konnte die Vorlage-Arbeitsmappe nicht ï¿½ffnen. Name:", vorl_dateiname
        raise "Konnte die Vorlage-Arbeitsmappe nicht ï¿½ffnen. (Name: #{vorl_dateiname})"
      end
      #xl.Visible = true
      #ExcelZugriff.als_vorderstes_fenster
      mappe.Windows(1).Visible = true
      mappe.Windows(1).Activate

      startzeit = Time.now
      trc_temp :mappe_erst_start
      begin
        vd.schreibe_excelref_e2(mappe)
      rescue
        trc_aktuellen_error "schreibe_excelref_e2: #{dateiname}"
        raise "Problem beim Erstellen der Vergleichs-Arbeitsmappe #{$!}"
      end
      trc_info :vor_delete
      xl.DisplayAlerts = false
      mappe.Sheets("Komp").Delete
      xl.DisplayAlerts = true
      stopzeit = Time.now
      mappe.save
      mappe
      #File.open("#{WORK_DIRNAME}/tmp.yaml","w") {|f| f.puts vd.to_yaml}
end

# Dieser Code wird momentan nich verwendet:
def dbase_in_leere_mappe(vd, dateiname)
  require 'wandler/dbf2exl'
      xl = ExcelZugriff.application
      begin
        mappe = xl.Workbooks.Open("#{ORIG_DIRNAME}/Vorlage.xls")
        begin
          mappe.SaveAs(dateiname)
        rescue WIN32OLERuntimeError
          raise "Konnte die Arbeitsmappe nicht als #{dateiname} speichern"
        end
      rescue WIN32OLERuntimeError
        raise "Konnte die Vorlage-Arbeitsmappe nicht ï¿½ffnen."
      end
      #xl.Visible = true
      ExcelZugriff.als_vorderstes_fenster
      mappe.Windows(1).Visible = true
      mappe.Windows(1).Activate

      startzeit = Time.now
      vd.schreibe_excel_e3(mappe)
      stopzeit = Time.now
      mappe.save
      File.open("#{WORK_DIRNAME}/tmp.yaml","w") {|f| f.puts vd.to_yaml}
      return mappe
end

###########

KOPIERVORL_KOMP_NR = 2
def kopiere_k2_blatt(mappe, neue_komp_nr)
  neue_komp_bez = "K#{neue_komp_nr}"
  replace_args = {"What"           => "_k#{KOPIERVORL_KOMP_NR}",
                  "Replacement"    => "_"+neue_komp_bez.downcase,
                  "LookAt"         => XLPart,
                  "MatchCase"      => false,
                  "MatchByte"      => false,
                  "SearchFormat"   => false,
                  "ReplaceFormat"  => false}

  blatt1 = mappe.Sheets("K#{KOPIERVORL_KOMP_NR}")
  blatt1.Copy("Before"=>mappe.Sheets(mappe.Sheets.Count))
  blatt = mappe.Sheets(mappe.Sheets.Count-1)
  trc_info :neu_blatt, [neue_komp_bez, blatt]
  blatt.Name = neue_komp_bez
  blatt.Activate
  excel = mappe.Application
  blatt.Cells.Replace(replace_args)
  trc_info :blatt_Replace_fertig, neue_komp_bez

  blatt.Range("KompNr").Value = neue_komp_nr

  eingaben = mappe.Sheets("Eingaben")
  eingaben.Activate
  ["Vorgaben1", "Vorgaben2"].each { |bez|
    eingaben.Range(bez+"_k#{KOPIERVORL_KOMP_NR}").Copy
    eingaben.Range(bez+"_Ende").Select
    excel.Selection.Insert(XLShiftToRight)
    neubereich = excel.Selection
    neubereich.Replace(replace_args)
    trc_info "excel.Selection.Address", neubereich.Address
    #neuer_name = excel.Names.Add("Name"=>bez+"_"+neue_komp_bez,
     #                            "RefersTo"=>neubereich.Address )
    #neuer_name.RefersToRange = Selection
  }
  # wird nicht gebraucht:
  #blatt.Names("rkw_k_alle").Delete rescue WIN32OLERuntimeError nil
end

def kopiere_k2_namen(mappe, neue_komp_nr)
  eingaben = mappe.Sheets("Eingaben")
  neue_komp_bez = "K#{neue_komp_nr}"
  excel = mappe.Application

  abstand = neue_komp_nr - KOPIERVORL_KOMP_NR

  excel.Names.each { |name_obj|
    neu_name = name_obj.Name
    if neu_name.sub!(/_k2$/i, "_"+neue_komp_bez)
      begin
        adresse = name_obj.RefersToRange.Address("External"=>true)
        case adresse
        when /Eingaben/
          neu_adresse = name_obj.RefersToRange.Offset(0,abstand).Address("External"=>true)
        when /\]K2/
          neu_adresse = adresse.sub(/\]K2/, ']'+neue_komp_bez)
        else
          trc_info :adresse_ignoriert, [name_obj.RefersTo, "-------", neu_name ]
          next
        end
        trc_temp :adresse__alt_neu, [name_obj.RefersTo, adresse, neu_adresse, neu_name ]
        excel.Names.Add("Name"     => neu_name,
                        "RefersTo" => "="+neu_adresse )
      rescue
        trc_aktuellen_error "probl bei altname=#{name_obj.Name rescue nil}, neuname=#{neu_name}"
        next
      end
    end
  }
  eingaben.Range("KN_k#{neue_komp_nr}").Value = neue_komp_bez
  eingaben.Range("KN2_k#{neue_komp_nr}").Value = neue_komp_bez

end


class DbfDat::Vd
  def vorgaben_rekursiv_eintragen(mappe)

    trc_info :anz_vts, vts.size
    
    fuer_alle_nach_zweitem = proc do |anweisung|
      vts.each_with_index do |vt, idx|
        blatt_nr = idx + 1
        vt.blatt_nr = blatt_nr
        if blatt_nr > KOPIERVORL_KOMP_NR
          anweisung.call blatt_nr  
        end
      end      
    end
    
    fuer_alle_nach_zweitem.call( proc{|blatt_nr| kopiere_k2_blatt(mappe, blatt_nr)} )
    fuer_alle_nach_zweitem.call( proc{|blatt_nr| kopiere_k2_namen(mappe, blatt_nr)} )
    
    
    ExcelZuordnungNG1.const_get("FORMELN").each do |name, formel|
      trc_info :formel, [name, self.instance_eval(&formel)]
      mappe.Names.Add("Name"     => name.to_s,
                      "RefersTo" => self.instance_eval(&formel) )
    end

    mappe.Application.DisplayAlerts = false
    begin
    # Das letzte Blett war die Klammer für Summen ï¿½ber alle Komponenten
      mappe.Sheets(mappe.Sheets.Count).Delete
    ensure
      mappe.Application.DisplayAlerts = true
    end

    super

    vp.vorgaben_rekursiv_eintragen(mappe)
    st.vorgaben_rekursiv_eintragen(mappe)

    vts.each_with_index {|vt, idx|
      blatt_nr = idx + 1
      if $kurze_liste then
        trc_temp :VERLAENGERUNG, blatt_nr
        vt.bringe_blatt_auf_laenge(mappe, blatt_nr)
      end
      vk = vt.vk
      vk.vorgaben_rekursiv_eintragen(mappe, blatt_nr)
      vt.vorgaben_rekursiv_eintragen(mappe, blatt_nr)
    }
  end

  def schreibe_excelref_e2(mappe)
#    sortierte_vks.each_with_index {|vk, idx|
 #     vk.schreibe_excelref_e2(mappe, idx+1)
  #  }
    vts.each_with_index {|vt, idx|
      blatt_nr = idx+1
      vt.vk.schreibe_excelref_e2(mappe, blatt_nr)
      blatt = mappe.Sheets("K#{blatt_nr}")
      vt.schreibe_excelref_e2(blatt, blatt_nr)
      trc_info "vd: vt geschrieben", [vt.komp, vt.vtnr]
    }
  end

private
  def vts
    if not @vts then
      @vts = vks.map { |vk|
        vk.vts.select { |vt|
          vt.tarmod != "B"
        }.to_a
      }.flatten
      @vts.each_with_index {|vt, i| vt.exl_komp_nr = i+1}
    end
    @vts
  end

end

class DbfDat::Vp
end

class DbfDat::St
end



class DbfDat::Vk

  def schreibe_excelref_e2(mappe, komp_nr)

    trc_info :beginne_vk_refschex, komp_nr

    mappe.Sheets("Komp").Copy("After"=>mappe.Sheets(mappe.Sheets.Count))
    blatt = mappe.Sheets(mappe.Sheets.Count)
    blatt.Name = "K#{komp_nr}"

    #blatt = mappe.Sheets.Item("K#{komp_nr}")
    blatt.Activate

    trc_info "schex-vk vt1-dauer=", vt1.n * 12 + vt1.nm
    if vt1.n * 12 + vt1.nm > 0 #Dividenden nur bei echter Laufzeit
      zeile = $erste_zeile
      trc_temp "schr-exl-ref_vk_vor_vvs", self.komp
      vvs.each {|vv|
        vv.schreibe_exl_zeile(blatt, zeile, bonus?)
        zeile += 1
      }
    end
    trc_info :ende_vk_refschex, komp_nr

  end

  def vorgaben_rekursiv_eintragen(mappe, komp_nr)
    trc_info :beginne_vk_nr, komp_nr
    @komp_nr = komp_nr
    super
  end

end


class DbfDat::Vt
  attr_accessor :blatt_nr

  def schreibe_bonus(blatt)
    trc_info :vt_schbonus__beg_zeit, beg_zeit
    trc_temp :vt_schbonus__diff, beg_zeit - vt1.beg_zeit
    rk = vt1.rk_nach_zeit(beg_zeit - vt1.beg_zeit)
    if rk
      rk.schr_bonus(blatt)
    else
      trc_info "vt_schbonus kein-rk für beg=", beg_zeit
    end
  end

  def exl_komp_nr=(nr)
    raise "Bug: Doppelte Zuweisung an Excel-KompNr (#{@exl_komp_nr})"if defined? @exl_komp_nr
    @exl_komp_nr = nr
  end
  attr_reader :exl_komp_nr


  def vorgaben_rekursiv_eintragen(mappe, komp_nr)
    trc_info :beginne_vt_nr_tabnam, [komp_nr, self.class.table_name]

    super

    #trc_temp :va, va
    if va
      va.vorgaben_rekursiv_eintragen(mappe, komp_nr)
    end
  end


  def bringe_blatt_auf_laenge(mappe, blatt_nr)
    blatt = mappe.Sheets("K#{blatt_nr}")
    blatt.Activate
    #$trace_stufe = :temp
    erste_allgemeine = blatt.Range("AllgemeineZeile")
    erste_allgemeine.Copy
    trc_temp :vor_schleife_blattnr=, blatt_nr
    rks.each_with_index {|rk, idx|
      nextzeile = erste_allgemeine.Offset(idx+1,0)
      nextzeile.Select
      blatt.Paste
    }
    trc_info :kopierschleife_fertig, blatt_nr
  end


  def schreibe_excelref_e2(blatt, komp_nr)

    trc_temp "schr-exl-ref_vt_anfang", [self.komp, self.vtnr]
    #trace :rk_arity, method(:rk).arity
    if tarmod == "B"
      #trace :schex_vtvk, vk
      trc_temp "schr-exl-ref_vt_vor-bonus", [self.komp, self.vtnr]
      schreibe_bonus(blatt) if vk.bonus?
    else
      #rks.each_with_index {|rk, idx| rk.schreibe_excelref_e2(blatt, idx) }
      zeile = $erste_zeile
      trc_temp "schr-exl-ref_vt_vor-rk", [self.komp, self.vtnr]
      rks.each {|rk|
        #rk.positions_init(blatt)
        rk.schreibe_exl_zeile(blatt, zeile)
        if rk.vb then
        #  rk.vb.schreibe_exl_zeile(blatt, zeile)
        end
        zeile += 1
      }
      zeile = $erste_zeile
      trc_temp "schr-exl-ref_vt_vor-vb", [self.komp, self.vtnr]
      vbs.each {|vb|
        vb.schreibe_exl_zeile(blatt, zeile)
        zeile += 1
      }
      zeile = $erste_zeile
      trc_temp "schr-exl-ref_vt_vor-vf", [self.komp, self.vtnr]
      vfs.each {|vf|
        vf.schreibe_exl_zeile(blatt, zeile)
        zeile += 1
      }

      schreibe_excelwert(blatt, "ZP", zp)  # "ZP_k#{komp_nr}"
      schreibe_excelwert(blatt, "BP", btr)
      schreibe_excelwert(blatt, "BR", br)
      schreibe_excelwert(blatt, "NP", gewbtrg)

    end
    trc_info "schr-exl-ref_vt_fertig", [self.komp, self.vtnr]
#+# Fehlt noch:
=begin
          // Und jetzt noch die Beitragsfreien Leistungen eintragen:
          FilterFirst( Vf_, KompFilter);
          Zeile := 2;
          repeat
            if Vf_RR.AsFloat > 0 then
              ZellwertA1[ BfrLeist_Spalte, Zeile ] := Vf_RR.AsFloat * Vf_RZW.AsFloat
            else if Vf_TSUM.AsFloat > 0 then
              ZellwertA1[ BfrLeist_Spalte, Zeile ] := Vf_TSUM.Value
            else
              ZellwertA1[ BfrLeist_Spalte, Zeile ] := Vf_ESUM.Value;

            Inc(Zeile);
          until not Vf_.FindNext;
=end
#    blatt.Range(blatt.Cells(2,                  4),
 #               blatt.Cells(2-1+val_array.size, 4-1+Rk.columns.size)).Value =
  #                  val_array
  end


end

class DbfDat::Va
end

class DbfDat::Rk
  def schr_bonus(blatt)
    rk.positions_init(blatt)
    rk.schr_tab_wert(:bonleihinzu, erster_positiver_wert(jren, tsum, esum) )
    rk.schr_tab_wert(:bonbtr, btr)
  end

  def schr_tab_wert(blatt, spaltenname, wert)
    spaltenname = spaltenname.to_s
    sp_nam_obj = exl_name_objekt(blatt, spaltenname)
    if sp_nam_obj
      spalten_nr = sp_nam_obj.ReferstoRange.Column
      excelwert_eintragen(blatt.Cells(exl_zeile, spalten_nr), wert)
    else
      trc_info :RK_wert_nicht_geschr
    end
  end

  def exl_zeile
    #idx = vt.rks.index(self)
    idx = vt_index
    if idx
      trc_temp "Rk-zeile idx=",idx
      $erste_zeile + idx
    else
      raise "Rk-zeile vtindex nicht gesetzt! vj=#{vj}"
      #trace "Rk-zeile self_nicht_in_rks!! vtnr=", vtnr
      nil
    end
  end


  def schreibe_exl_zeile(blatt, zeile)
      #if @exl_zeile then
      #  raise "Bug: Es lï¿½uft schon ein Zeilenschreibvorgang (zeile=#{zeile}, tarif=#{vt.tarif})"
      #end
      #@exl_zeile = zeile
      #@blatt = blatt

      excelwert_eintragen(blatt.Range("A#{zeile}:B#{zeile}"), [vj, vm]) # Reinschreiben als Range, weil Matrixformel in der Tabelle steht:

      if gezdkk != 0
        schr_tab_wert(blatt, :vhgb, gezdkk)
      else
        schr_tab_wert(blatt, :vhgb, rmdkk)
      end

      schr_tab_wert(blatt, :vias,      netdrk)
      schr_tab_wert(blatt, :rkw,       rkw)
      schr_tab_wert(blatt, :aktivw,    aktivw)

      schr_tab_wert(blatt, :RmAbsk,    rmak)   rescue nil
      schr_tab_wert(blatt, :LaufAZ,    laufaz) rescue nil
      schr_tab_wert(blatt, :VwkPr,     vwkpr)  rescue nil
      schr_tab_wert(blatt, :RatzusVwk, rzvk)   rescue nil
      schr_tab_wert(blatt, :StornoAb,  stornoab) rescue nil


=begin
          Zeile := 1;
          if FilterFirst( Vb_, KompFilter, False) then
          repeat
          //  Handstand!!! fange eine Zeile frï¿½her an, damit Leistung zweimal geschrieben wird (doch nicht veraltet: 2005-Jum-10 Svs)
            if Zeile >= 2 then // d.h. beim ersten Mal noch nicht.
              ZellwertA1[ Btr_Spalte, Zeile] := Vb_BTRVJ.Value;


            Inc(Zeile);

            // Dies ist also das einzige, was beim ersten Mal ausgefï¿½hrt wird:
            if Vb_JRENVJ.Value > 0 then
              ZellwertA1[ LeistSaldiert_Spalte, Zeile ] := Vb_JRENVJ.Value
            else if Vb_TSUMVJ.Value > 0 then
              ZellwertA1[ LeistSaldiert_Spalte, Zeile ] := Vb_TSUMVJ.Value
            else
              ZellwertA1[ LeistSaldiert_Spalte, Zeile ] := Vb_ESUMVJ.Value;
            // damit wird effektiv der erste Leistungswert in zwei ï¿½bereinaderliegende Zellen geschrieben.

            if (Zeile >= 3) then // d.h. beim ersten Mal noch nicht.
              if not Vb_.FindNext then BREAK;

          until false; //and
=end
    #@exl_zeile = nil
  end

end

class DbfDat::Vb
  def schr_tab_wert(blatt, zeile, spaltenname, wert)
    spaltenname = spaltenname.to_s
    sp_nam_obj = exl_name_objekt(blatt, spaltenname)
    if sp_nam_obj
      spalten_nr = sp_nam_obj.ReferstoRange.Column
      excelwert_eintragen(blatt.Cells(zeile, spalten_nr), wert)
    else
      trc_hinweis :vb_wert_nicht_geschr
    end
  end

  def schreibe_exl_zeile(blatt, zeile)
    trc_temp "schr-exl-ref_vb_schr-zeile", [self.komp, self.vj, self.vm, zeile]
    schr_tab_wert(blatt, zeile, :btr, btrvj)
  end

end

class DbfDat::Vf
  def schr_tab_wert(blatt, zeile, spaltenname, wert)
    spaltenname = spaltenname.to_s
    sp_nam_obj = exl_name_objekt(blatt, spaltenname)
    if sp_nam_obj
      spalten_nr = sp_nam_obj.ReferstoRange.Column
      excelwert_eintragen(blatt.Cells(zeile, spalten_nr), wert)
    else
      trc_hinweis :vb_wert_nicht_geschr
    end
  end

  def schreibe_exl_zeile(blatt, zeile)
    trc_temp "schr-exl-ref_vf_schr-zeile", [self.komp, self.vj, self.vm, zeile]

    leist = if rr > 0 then
      rr
    elsif tsum > 0 then
      tsum
    else
      esum
    end
    schr_tab_wert(blatt, zeile, :BfrLeist, leist)
  end

end

class DbfDat::Vv
  def schr_tab_wert(spaltenname, wert)
    spaltenname = spaltenname.to_s
    sp_nam_obj = exl_name_objekt(@blatt, spaltenname)
    if sp_nam_obj
      spalten_nr = sp_nam_obj.ReferstoRange.Column
      excelwert_eintragen(@blatt.Cells(@exl_zeile, spalten_nr), wert)
    else
      trc_info "Vv_wert_nicht_geschr (#{spaltenname})", wert
    end
  end

  def schreibe_exl_zeile(blatt, zeile, mit_bonus=false)

      @blatt = blatt
      @exl_zeile = zeile

      schr_tab_wert(:ZinsDivGes, zindiv)
      schr_tab_wert(:RisDivGes,  risdiv)
      schr_tab_wert(:BeiDivGes,  beidiv)
      schr_tab_wert(:GruDivGes,  grdiv)
      schr_tab_wert(:adg,        adg)
      schr_tab_wert(:sdfond,     sdfond)
      #trc_temp "sdrueck", "sdrï¿½ck"
      schr_tab_wert("sdrï¿½ck",    sdrueck)  # #+# #!# ### Testen!!


      if mit_bonus then
        schr_tab_wert(:BonLeiKumul, erster_positiver_wert(bonjr, bonsum))
        schr_tab_wert(:BonDkk,    bondkk)
      end
  end

end

class DbfDat::VObjektBasis
  def vorgaben_rekursiv_eintragen(mappe, komp_nr=nil)
    exl_vorgaben_eintragen(mappe, komp_nr)
  end

  def exl_werte
    trc_temp :table_data, self.class.tabelle.data
    erg = {}
    prefix = self.class.name.split("::").last.upcase
    ExcelZuordnungNG1.const_get(prefix+"_zuord").each do |exlname, exl_felddef|
      #exl_felddef = (self.class)::FELDER_EXL[exlname.to_sym]
      wert =
        case exl_felddef
        when :nix  then next
        when Proc  then
          begin
            self.instance_eval(&exl_felddef)
          rescue
            trc_aktuellen_error :felddef_eval, 10
            $ats.konsole.meldung "Wert für Zelle:#{exlname} konnte nicht berechnet werden. Fehler=#{$!}"
          end
        else
          begin
            self.attributes[exl_felddef.to_s]
          rescue
            trc_aktuellen_error :felddef_ausfeld, 8
            $ats.konsole.meldung "Zelle:#{exlname} ist mit #{self.class.name.downcase}.#{exl_felddef} verknï¿½pft, das Feld wurde aber nicht in der Tabelle gefunden"
          end
        end

      trc_temp :exlname_wert_def , [exlname, wert, exl_felddef]
      erg[exlname] = wert
    end
    erg
  end

  def exl_vorgaben_eintragen(mappe, komp_nr=nil)
    eingaben = mappe.Sheets("Eingaben")
    eingaben.Activate
    exl_werte.each do |exlname, wert|
      exlname = "#{exlname}_k#{komp_nr}" if komp_nr
      begin
        zelle = exl_range(eingaben, exlname.to_s)
        if not zelle then
          $ats.konsole.meldung "Name:#{exlname} ist keinem Bereich in der Vorlage zugeordnet"
        end
        excelwert_eintragen(zelle, wert)
      rescue
        trc_aktuellen_error "Range #{exlname}"
        $ats.konsole.meldung "Fehler beim Schreiben in Zelle:#{exlname}, Meldung=#{$!}"
      end
    end
  end

  def exl_range(blatt, name)
    namens_objekt = exl_name_objekt(blatt, name)
    namens_objekt.RefersToRange if namens_objekt
  end

  def exl_name_objekt(blatt, name)
    begin
      blatt.Parent.Names.Item(name)
    rescue WIN32OLERuntimeError
      #trace :versuche_imblatt, name
      begin
        blatt.Names.Item(name)
      rescue WIN32OLERuntimeError
        nil
      end
    end
  end

  def schreibe_excelwert(blatt, zellname, wert) #name, wert)
    name = zellname.to_s
    #trace :schexcelwert, zellname
    if exl_name_objekt(blatt, name)
      excelwert_eintragen(blatt.Range(name), wert)
    else
      trc_hinweis "nichtgeschrieben", zellname
    end
  end

  def excelwert_eintragen(zelle, wert)
    if zelle then
	    zelle.Value = wert
	    zelle.Interior.ColorIndex =
	    #  7 # Grï¿½n
	      42 #  "Aquamarin"
    else
      trc_hinweis :zelle_nil__wert=, wert
    end
  end

  def schr_tab_wert(blatt, zeile, spaltenname, wert)
    spaltenname = spaltenname.to_s
    sp_nam_obj = exl_name_objekt(blatt, spaltenname)
    if sp_nam_obj
      spalten_nr = sp_nam_obj.ReferstoRange.Column
      excelwert_eintragen(blatt.Cells(zeile, spalten_nr), wert)
    else
      trc_hinweis :vb_wert_nicht_geschr
    end
  end

  def schreibe_exl_zeile(blatt, zeile)
    trc_temp "schr-exl-ref_vb_schr-zeile", [self.komp, self.vj, self.vm, zeile]
    schr_tab_wert(blatt, zeile, :btr, btrvj)
  end

protected
  def eigene_werte_eintragen(komp_nr)
    #zu lang: trace :attributes, [self, attributes]

    eingaben = @mappe.Sheets("Eingaben")
    eingaben.Activate
    #attributes.each { |feldsym, wert|
    (self.class)::FELDER_EXL.each { |feldsym, exl_felddef|
      #exl_felddef = (self.class)::FELDER_EXL[feldsym.to_sym]
      exl_feldname =
        case exl_felddef
          when :nix  then next
          when nil   then feldsym.to_s
          when Hash  then exl_felddef.keys.first.to_s
          else            exl_felddef
        end

      trc_temp :exl_feldname , [feldsym, exl_feldname ]
      begin
        exl_feldname += "_k#{komp_nr}" if komp_nr
        #trace :komplname, exl_feldname
        zelle = eingaben.Range(exl_feldname)
      rescue
        next
      end
      wert = attributes[feldsym.to_s]
      wert = exl_felddef.values.first.call(wert) if exl_felddef.is_a?(Hash)
      trc_info :wert, wert
      excelwert_eintragen(zelle, wert)
    }
    #neuer_bereich(blatt, feldsyms)
  end

end

end # __FILE__ == $0

end # if not method_defined? :dbf2exl_vsnr 
