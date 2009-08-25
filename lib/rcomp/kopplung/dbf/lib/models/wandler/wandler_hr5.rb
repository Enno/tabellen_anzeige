
if not defined? WandlerAllgemein then

require File.dirname(__FILE__) + '/../../schmiedebasis'

require 'models/system_fkt/dbase_zugriff'
require 'models/generali'
require 'models/system_fkt/excel_zugriff'

require 'models/wandler/excel_zuordnung'
  
require 'models/wandler/wandler_allgemein'


class Wandler_Hr5 < WandlerVertrZuMappe
  
  def initialize(quellort_pfad, ziel_pfad, opts = {})    
    super
    
    @opts = {:mappen_offenhalten => false}.merge(@opts)
    
    trc_info "waart_str, ziel_pfad", [@waart_str, quellort_pfad, ziel_pfad]
    begin
      #vsnr_pfad = vsnr_pfad.gsub(/[<>]/,"")
      ziel_pfad = ziel_pfad.gsub(/[<>]/,"")
      trc_info :dbf2exlvsnr_ziel, ziel_pfad
  
      DbfDat::oeffnen(quellort_pfad, :s)
  
  
    ensure
    end
    
  end
  
  def schlieszen
    DbfDat::schliessen
    
  end

  def wandle_vertrag(vsnr, ziel_name)
    startzeit = Time.now
    
    vsnr = File.basename(vsnr)
    trc_temp :erz_excel__vsnr=, vsnr

    vd = DbfDat::Vd.find( :first, "vsnr" => vsnr)
    if !vd then
      trc_fehler "vd nicht gefunden", "vsnr='#{vsnr}', pfad='#{@quellort_pfad}'"
      raise "dbf-Vertrag (vsnr='#{vsnr}') nicht in '#{@quellort_pfad}' gefunden!"
    end
    trc_info :vd_found, vd.vsnr

    ExcelZugriff.sichtbar = KONFIG.opts[:MappenAnzeigenBeiBatchjobs]
    
    ma_wa = MappenWandlung.new(self, vd, ziel_name)
    ma_wa.erzeuge_zieldatei
    ma_wa.erzeuge_refdatei
    
    stopzeit = Time.now
    dauer_gespeichert = stopzeit - startzeit    
    trc_temp :wandel_dauer, dauer_gespeichert  
    
    dauer_gespeichert    
  end

end

class WandelError < RuntimeError  
end

EXCEL_REPLACE_PARAMS_STD = {
                "What"           => "nix",
                "Replacement"    => "nix",
                "LookAt"         => XLPart,
                "MatchCase"      => false,
                "MatchByte"      => false,
                "SearchFormat"   => false,
                "ReplaceFormat"  => false
}.freeze
                  
class BlattFamilie < Struct.new(:prefix, 
                                :format, 
                                :allgemeiner_name, 
                                :gegebene_namen, 
                                :zu_loeschender_name,
                                :vo_typ)
  def initialize(prefix, namen)
      
      self.prefix = prefix
      
      self.zu_loeschender_name = namen.select{|n| n=~/^[^#\d]+9+$/}.sort.last
      
      self.gegebene_namen = namen - [self.zu_loeschender_name]
      
      self.allgemeiner_name = self.gegebene_namen.sort.last
      if not self.allgemeiner_name then
        next if namen.size <= 1
        next if prefix.size <= 1 and namen.size <= 2
        raise "Für die BlattFamilie mit '#{prefix}...' fehlt das allgemeine Blatt"
      end
      stellenanz_allg = self.allgemeiner_name.size - prefix.size

      konstante_breite = if self.zu_loeschender_name then
        stellenanz_max  = self.zu_loeschender_name.size - prefix.size
        0 == (stellenanz_max - stellenanz_allg)
      else
        "0" == self.allgemeiner_name[prefix.size, 1]
      end
      
      ist_numerisch = (self.allgemeiner_name[prefix.size..-1] =~ /[#\d]+/)
      
      self.format = if konstante_breite and ist_numerisch then
        "#{prefix}%0#{stellenanz_allg}d"
      else
        "#{prefix}%s"
      end      

      self.vo_typ = case prefix
                    when /^(vt|t_?$|rk|vb)/ then :vt
                    when /^(vk|k_?$|vv)/ then :vk
                    end
      
  end
  
  def blatt_name_fuer(bezeichnung)  
    format % bezeichnung
  rescue ArgumentError
    raise WandelError, "Bezeichnung '#{bezeichnung}' passt nicht ins Format (#{format}) der Blattnamen"
  end
  
  def blatt_existenz_sicherstellen(mappe, bezeichnung)    
    #@kopiervorl_blatt_bez = satz_id_zu_blatt_bez(KOPIERVORL_KOMP_NR)
    neuer_blatt_name = blatt_name_fuer bezeichnung.to_s
    
    return if gegebene_namen.include?(neuer_blatt_name)

    replace_args = EXCEL_REPLACE_PARAMS_STD.dup.update(
                    "What"           => "_"+allgemeiner_name.downcase,
                    "Replacement"    => "_"+neuer_blatt_name.downcase
    )
    replace_args_plus = EXCEL_REPLACE_PARAMS_STD.dup.update(
                    "What"           => allgemeiner_name + ".",
                    "Replacement"    => neuer_blatt_name + "."  
    )                 

    allgemein_blatt = mappe.Sheets(allgemeiner_name)
    ExcelZugriff.ohne_displayalerts do
      allgemein_blatt.Copy("Before"=>mappe.Sheets(mappe.Sheets.Count))
    end
    neues_blatt = mappe.Sheets(mappe.Sheets.Count-1)
    trc_info :neu_blatt, [neuer_blatt_name, neues_blatt]
    neues_blatt.Name = neuer_blatt_name
    neues_blatt.Activate
    
    ExcelZugriff.ohne_displayalerts do
      neues_blatt.Cells.Replace(replace_args)
      neues_blatt.Cells.Replace(replace_args_plus)    
    end
    trc_info :blatt_Replace_fertig, neuer_blatt_name
    
    #neues_blatt.Range("AktSatzBez").Value = bezeichnung
  end  
##########
end


class MappenWandlung 
  
  attr_reader :wandler, :vd, :vsnr, :zieldateiname
  attr_reader :excel, :mappe, :refmappe
  
  KOPIERVORL_KOMP_NR = 1
  KOPIERVORL_ZEIL_NR = 1
  
  def initialize(wandler, vd, ziel_name)
    @wandler = wandler
    @vd = vd    
    @vsnr = @vd.vsnr
    @quell_pfad = @wandler.quellort_pfad + "/" + vsnr
    @zielbasisname = ziel_name
    
    @zieldateiname = wandler.ziel_pfad + "/" + @zielbasisname
    
    @refdateiname = wandler.ziel_pfad + '/Referenzdaten/'+ExcelZugriff::STD_REFPREFIX + @zielbasisname
    
    Phi::force_dirs(File.dirname(@refdateiname))
    trc_info :refdateiname, @refdateiname
    
    @vertikal_ausdehnen = true
  end
  
  def schreibe_info
    if true then      
      kurzinfo = begin
        ort_persist = OrtPersistenz.new(File.dirname(@wandler.quellort_pfad))
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
        "DateiQuelle" => @quell_pfad
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
  
  def erzeuge_zieldatei
    begin
      trc_info :HR5_zieldatei_start, zieldateiname
      
      vorlage_initialisieren
      sammle_blattfamilien
      
      zielmappe_ausdehnen      
      ziel_mit_vorgaben_fuellen
      
      trc_info :HR5_zieldatei_fertig, zieldateiname
    rescue Exception => exception
      if ExcelError === exception then
        mappe.Close(false) rescue nil
        trc_aktuellen_error "Mappe steht im Weg beim Erstellen der Vorgaben-Arbeitsmappe", 3      
        raise "Kann Vorgaben-Mappe nicht erstellen: #{exception}"
      else
        trc_aktuellen_error "Problem beim Erstellen der Vorgaben-Arbeitsmappe"      
        raise "Problem beim Erstellen der Vorgaben-Arbeitsmappe (\"#{
                                     $!.to_s.split("\n")[0][0..70].strip}\")"
        
      end
    end
    mappe.save

    # mappen_endbehandlung 
  end

  def vorlage_initialisieren
    @mappe = ExcelZugriff.excel_vorlage_oeffnen(wandler.waart_obj)     
    @excel = mappe.Application

    ExcelZugriff.speichere_kontrolliert(mappe, zieldateiname, KONFIG.opts[:WennBeiExcelErstellungMappeExistiert])

    schreibe_info    
  end

  def sammle_blattfamilien
    worksheets = @mappe.Sheets
    worksheets.extend Enumerable
    
    alle_blattnamen = worksheets.map { |blatt| blatt.Name }
    
    eingaben_blatt_name = alle_blattnamen.find do |bn|
      bn =~ /(Vor|Ein)gaben/
    end
    raise WandelError, "Exstar-Vorlage besitzt kein Blatt 'Eingaben'" unless eingaben_blatt_name
    @eingaben_blatt = mappe.Sheets(eingaben_blatt_name)
    alle_blattnamen.delete(eingaben_blatt_name)
    blatt_familien_arten = [
      /^([^_#\d]+)[#\d]+$/,
      /^(.+_)[^_]+$/,
      /^()[^.]+\.[^.]+$/
    ]
    eine_der_arten = Regexp.union(*blatt_familien_arten)
    interessante_blattnamen = alle_blattnamen.grep(eine_der_arten)
    
    blattnamen_nach_arten = {}
    interessante_blattnamen.each do |name|
      anfang = blatt_familien_arten.each do |art_regexp|
        if name.match(art_regexp) then
          break $1
        end
      end
      raise "bug: nix gefunden (#{anfang})" if anfang.is_a? Array
      #name.match(/^([^#\d_]+|.+_)[^_]+$/)[1]      
      (blattnamen_nach_arten[anfang] ||= []) << name
    end

    @blatt_familien = {}
    blattnamen_nach_arten.each do |anfang, namen|
      @blatt_familien[anfang] = BlattFamilie.new(anfang, namen)
    end

    raise "keine Blätter zur Erzeugung in Vorlage vorhanden" if @blatt_familien.empty?
    
    sortierte_familien = @blatt_familien.values.sort_by { |bf| bf.prefix }
    
    @haupt_blattfamilie_tsatz = sortierte_familien.select{|bf| bf.vo_typ == :vt}.last    
    @haupt_blattfamilie_tsatz ||= sortierte_familien.last    
    
    
    @benannter_eingaben_bereich = {}
    @haupt_blattfamilie_tsatz.gegebene_namen.sort.each do |blatt_name|
      bereich = (@eingaben_blatt.Range(blatt_name+".") rescue nil)
      @benannter_eingaben_bereich[blatt_name] = bereich if bereich
    end
    
    @wachsende_vorgaben_bereichsnamen = [
      @haupt_blattfamilie_tsatz.allgemeiner_name + ".",
      "Vorgaben1_k2", 
      "Vorgaben2_k2"
    ].select { |name|  (@mappe.Application.Names(name) rescue false) }
    
    @blatt_familien
  end  
  
#private 
  def zielmappe_ausdehnen

    trc_info :anz_vts, vd.normale_vts.size
    
    @blatt_familien.each do |prefix, blatt_sorte|
      vobjekte = case blatt_sorte.vo_typ
      when :vt then vd.normale_vts
      when :vk then vd.vks
      else raise WandelError, "Konnte für Präfix '#{prefix}' den Typ der Blatt-Familie nicht zuordnen"
      end
      vobjekte.each_with_index do |vobjekt, idx|
        bez = case vobjekt
        when DbfDat::Vt then idx+1
        when DbfDat::Vk then vobjekt.komp
        end
        blatt_sorte.blatt_existenz_sicherstellen(mappe, bez) 
      end
    end
    #fuer_alle_nach_zweitem_vtsatz() {|blatt_nr| kopiere_k2_blatt(blatt_nr) }
    abstand = 0
    vd.normale_vts.each_with_index do |vt, idx| 
      blatt_nr = idx + 1
      neuer_blatt_name = @haupt_blattfamilie_tsatz.blatt_name_fuer blatt_nr
      if not @haupt_blattfamilie_tsatz.gegebene_namen.include?(neuer_blatt_name) then
        abstand += 1
        vorgaben_bereich_sicherstellen(blatt_nr, abstand)
        namen_fuer_neues_blatt_erzeugen(blatt_nr, abstand)
      end
    end
    #fuer_alle_nach_zweitem_vtsatz() {|blatt_nr| kopiere_k2_namen(blatt_nr) }
    
    zielformeln_eintragen
    
    ExcelZugriff.ohne_displayalerts do
      # Das letzte Blett war nur die Klammer für Summen über alle Komponenten
      mappe.Sheets(mappe.Sheets.Count).Delete
    end
  end
    
  def xx_fuer_alle_nach_zweitem_vtsatz( &anweisung )
    vd.normale_vts.each_with_index do |vt, idx|
      blatt_nr = idx + 1
      #vt.blatt_nr = blatt_nr     
      anweisung.call(blatt_nr) if blatt_nr > KOPIERVORL_KOMP_NR
    end      
  end
  
  def satz_id_zu_blatt_bez(tsatz_bezeichnung)    
    #"#{@blattprefix}#{'%02d' % satz_id}"
    @haupt_blattfamilie_tsatz.blatt_name_fuer tsatz_bezeichnung    
  end
  
  
  def vorgaben_bereich_sicherstellen(tsatz_bezeichnung, abstand)
    neuer_blatt_name = @haupt_blattfamilie_tsatz.blatt_name_fuer tsatz_bezeichnung
    #return if @haupt_blattfamilie_tsatz.gegebene_namen.include?(neuer_blatt_name)
    allg_blatt_name = @haupt_blattfamilie_tsatz.allgemeiner_name
    
    @eingaben_blatt.Activate

    replace_args = EXCEL_REPLACE_PARAMS_STD.dup.update(
                    "What"           => "_"+allg_blatt_name.downcase,
                    "Replacement"    => "_"+neuer_blatt_name.downcase
    )

    richtung = (@vertikal_ausdehnen ? XLShiftDown : XLShiftToRight)
    @wachsende_vorgaben_bereichsnamen.each do |name|
      begin
        orig_bereich = @eingaben_blatt.Range(name)
        ziel_bereich = verschiebe_eing(orig_bereich, abstand)
        orig_bereich.Copy
        ziel_bereich.Select
        excel.Selection.Insert(richtung)
#        abstand = @vertikal_ausdehnen ? 
#          neu_bereich.Row    - orig_bereich.Row :
#          neu_bereich.Column - orig_bereich.Column
        neu_bereich = excel.Selection
        ExcelZugriff.ohne_displayalerts do
          neu_bereich.Replace(replace_args)
        end
        trc_info "excel.Selection.Address", neu_bereich.Address
        #neuer_name = excel.Names.Add("Name"=>bez+"_"+neue_komp_bez,
         #                            "RefersTo"=>neubereich.Address )
        #neuer_name.RefersToRange = Selection
                    
      rescue WIN32OLERuntimeError
        trc_aktuellen_error("Eingabebereich #{name} nicht kopierbar", 8)
      end  
    end
  end
  
  def verschiebe_eing(bereich, abstand)
    bereich.Offset(*eingabenblatt_richtungs_vektor(abstand))
  end
  
  def eingabenblatt_richtungs_vektor(abstand) # zeile, spalte
    if @vertikal_ausdehnen then 
      [abstand,0]
    else
      [0,abstand]
    end
  end

  def namen_fuer_neues_blatt_erzeugen(tsatz_bezeichnung, abstand)    
    #@kopiervorl_blatt_bez = satz_id_zu_blatt_bez(KOPIERVORL_KOMP_NR)
    neuer_blatt_name = @haupt_blattfamilie_tsatz.blatt_name_fuer tsatz_bezeichnung
    #return if @haupt_blattfamilie_tsatz.gegebene_namen.include?(neuer_blatt_name)
    allg_blatt_name = @haupt_blattfamilie_tsatz.allgemeiner_name
    #neue_blatt_bez = satz_id_zu_blatt_bez(neue_komp_nr)
    #neuer_komp_suffix = "K#{neue_komp_nr}"
    

    #richtungs_vektor = eingabenblatt_richtungs_vektor(abstand)
    eingaben_blatt_name = @eingaben_blatt.Name
    
    excel.Names.each do |name_obj|      
      neu_name = name_obj.Name
      if neu_name.sub!(/_#{allg_blatt_name}$/i, "_"+neuer_blatt_name)
        excel_namen_generieren(neu_name) do
          adresse = name_obj.RefersToRange.Address("External"=>true)
          case adresse
          when /#{eingaben_blatt_name}/
            neu_adresse = verschiebe_eing(name_obj.RefersToRange, abstand).Address("External"=>true)
          when /\]#{allg_blatt_name}/
            neu_adresse = adresse.sub(/\]#{allg_blatt_name}/, ']'+neuer_blatt_name)
          else
            trc_info :adresse_ignoriert, [name_obj.RefersTo, "-------", neu_name ]
            next
          end
          neu_adresse
        end
        
      end
    end
    
    # #*# oben mit einbauen
    if not @benannter_eingaben_bereich.empty? then
      neuer_name = neuer_blatt_name + "."
      
      erg = excel_namen_generieren(neuer_name) do      
        orig_bereich = @benannter_eingaben_bereich[allg_blatt_name]
        neu_bereich = verschiebe_eing(orig_bereich, abstand)
        @benannter_eingaben_bereich[neuer_blatt_name] = neu_bereich
        neu_bereich.Address("External"=>true)
      end
              
    end
    
    #@eingaben_blatt.Range("KN_k#{neue_komp_nr}").Value = neue_blatt_bez rescue nil
    #@eingaben_blatt.Range("KN2_k#{neue_komp_nr}").Value = neue_blatt_bez rescue nil
  end
  
  def excel_namen_generieren(neuer_name)
    neue_adresse = yield
    excel.Names.Add("Name"     => neuer_name,
                                "RefersTo" => "="+neue_adresse )      
    trc_temp :adresse_neu, [neuer_name, neue_adresse]
    neuer_name
  rescue
    trc_aktuellen_error "probl beim Erstellen von neuname=#{neuer_name}"  
    nil
  end
  
  def zielformeln_eintragen
    ExcelZuordnungHR5.const_get("FORMELN").each do |name, formel|
      excel_string = vd.instance_eval(&formel)
      trc_info :formel, [name, excel_string]
      ExcelZugriff.ohne_displayalerts do
        mappe.Names.Add("Name"     => name.to_s,
                        "RefersTo" => excel_string)
      end
    end
  end
  
  
    
  def ziel_mit_vorgaben_fuellen
    vorgabewerte_eintragen_fuer(vd)
    vorgabewerte_eintragen_fuer(vd.vp)
    vorgabewerte_eintragen_fuer(vd.st)

    vd.normale_vts.each_with_index do |vt, idx|
      blatt_nr = idx + 1
      blatt_name = @haupt_blattfamilie_tsatz.blatt_name_fuer(blatt_nr)
      bringe_blatt_auf_laenge(vt, blatt_name)      
      vk = vt.vk
      vorgabewerte_eintragen_fuer(vk, blatt_nr)
      vorgabewerte_eintragen_fuer(vt, blatt_nr)
      vorgabewerte_eintragen_fuer(vt.va, blatt_nr) if vt.va
    end
  end

  def einzutragende_werte_fuer(vobjekt)
    trc_temp :classname, vobjekt.class.name #.tabelle.data
    erg = {}
    prefix = vobjekt.class.name.split("::").last.upcase
    zuordnung = ExcelZuordnungHR5.const_get(prefix+"_zuord")
    zuordnung.each do |exlname, exl_felddef|
      #exl_felddef = (self.class)::FELDER_EXL[exlname.to_sym]
      wert = case exl_felddef
      when :nix  then 
        next
      when Proc  then
        begin
          vobjekt.instance_eval(&exl_felddef)
        rescue
          trc_aktuellen_error :felddef_eval, 10
          $ats.konsole.meldung "Wert für Zelle:#{exlname} konnte nicht berechnet werden. Fehler=#{$!}"
        end
      else
        begin
          vobjekt.attributes[exl_felddef.to_s]
        rescue
          trc_aktuellen_error :felddef_ausfeld, 8
          $ats.konsole.meldung "Zelle:#{exlname} ist mit #{vobjekt.class.name.downcase}.#{exl_felddef} verknüpft, das Feld wurde aber nicht in der Tabelle gefunden"
        end
      end

      trc_temp :exlname_wert_def , [exlname, wert, exl_felddef]
      erg[exlname] = wert
    end
    erg
  end
  
  def vorgabewerte_eintragen_fuer(vobjekt, komp_nr=nil)  
    @eingaben_blatt.Activate
    einzutragende_werte_fuer(vobjekt).each do |exlname, wert|
      exlname = exlname.to_s
      # #*# Hack für alte Vorlage-Versionsn
      exlname += "_k#{komp_nr}" if komp_nr and @benannter_eingaben_bereich.empty?
      begin
        zelle_oder_vektor = exl_range(@eingaben_blatt, exlname)
        if not zelle_oder_vektor then
          $ats.konsole.meldung "Name:#{exlname} ist keinem Bereich in der Vorlage zugeordnet"
        else
          zelle = if komp_nr and not @benannter_eingaben_bereich.empty? then
            blatt_name = @haupt_blattfamilie_tsatz.blatt_name_fuer(komp_nr)
            @mappe.Application.Intersect(@benannter_eingaben_bereich[blatt_name], zelle_oder_vektor)
          else
            zelle_oder_vektor
          end
          excelwert_eintragen(zelle, wert)
        end
      rescue
        trc_aktuellen_error "Range #{exlname}"
        $ats.konsole.meldung "Fehler beim Schreiben in Zelle:#{exlname}, Meldung=#{$!}"
      end
    end
    
  end

  def bringe_blatt_auf_laenge(vt, blatt_name)
    blatt = mappe.Sheets(blatt_name)
    blatt.Activate
    
    erste_allgemeine = blatt.Range("AllgemeineZeile")
    
    trc_temp :vor_schleife_blattnr=, blatt_name
    vt.rks.each_with_index do |rk, idx|
      erste_allgemeine.Copy
      nextzeile = erste_allgemeine.Offset(idx+1,0)
      nextzeile.Select
      
      #excel.Selection.Insert(XLShiftDown)
      blatt.Paste
    end
    trc_info :kopierschleife_fertig, blatt_name
  end

  def exl_range(blatt, name)
    namens_objekt = exl_name_objekt(blatt, name)
    if namens_objekt then 
      begin
        namens_objekt.RefersToRange 
      rescue WIN32OLERuntimeError
        bezug = (namens_objekt.RefersToLocal rescue nil)
        raise "Ungültiger Bezug (#{bezug}) im Namen #{name}"
      end
    end          
  end

  def exl_name_objekt(blatt, name)
    begin
      blatt.Parent.Names.Item(name)
    rescue WIN32OLERuntimeError
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


###########################################################  
public 
  def erzeuge_refdatei
    if @refdateiname == @zieldateiname then
      raise "Bug: Referenz-Name darf nicht identisch zum Original sein: #{@refdateiname}"
    end
    
    if not @mappe then
      @mappe = excel_zugriff..oeffne_oder_aktiviere(@zieldateiname, :raise)
    end
    ExcelZugriff.speichere_kontrolliert(@mappe, @refdateiname, KONFIG.opts[:WennBeiExcelErstellungMappeExistiert])
    
    refmappe_mit_werten_fuellen
    
    ExcelZugriff.nurnoch_werte_behalten(mappe)
    
    mappe.save
    
    mappen_endbehandlung 
  end
  
  def refmappe_mit_werten_fuellen
    vd.normale_vts.each_with_index do |vt, vt_idx|
      satz_nr = vt_idx + 1
      blatt_bez = satz_id_zu_blatt_bez(satz_nr)
      blatt = mappe.Sheets(blatt_bez)
      blatt.Activate
      erste_zeile_nr = blatt.Range("ErsteZeile").Row
      vt.rks.each_with_index do |rk, rk_zeile|
        refwerte_eintragen_fuer( rk,   blatt,  erste_zeile_nr + rk_zeile)
      end
      trc_info "vt geschrieben, [satz_id,komp,vtnr]=", [satz_nr, vt.komp, vt.vtnr]
    end
  end
  
  def refwerte_eintragen_fuer(vobjekt, blatt, zeilen_nr)    
    ganze_aktuelle_zeile = ( blatt.Range("$#{zeilen_nr}:$#{zeilen_nr}") if zeilen_nr )
    trc_temp "zeilnr, gaz_adr", [zeilen_nr, ganze_aktuelle_zeile.Address]
    einzutragende_werte_fuer(vobjekt).each do |exlname, wert|
      begin
        zelle_oder_vektor = exl_range(blatt, exlname.to_s)
        if not zelle_oder_vektor then
          $ats.konsole.meldung "Name:#{exlname} ist keinem Ergebnis-Bereich in der Vorlage zugeordnet"
        else
          zelle = if ganze_aktuelle_zeile then
            @mappe.Application.Intersect(ganze_aktuelle_zeile, zelle_oder_vektor)
          else
            zelle_oder_vektor
          end
          excelwert_eintragen(zelle, wert)
        end
      rescue
        trc_aktuellen_error "Range #{exlname}"
        $ats.konsole.meldung "Fehler beim Schreiben in Zelle:#{exlname}, Meldung=#{$!}"
      end
    end
  end
  


###########################################################  
public
  def mappen_endbehandlung 
    if mappe and not wandler.opts[:mappen_offenhalten]
      mappe.Close 
      @mappe = nil
    else
      mappe
    end
  end
  
end

class DbfDat::Vd

  def normale_vts
    if not @vts then
      @vts = vks.map { |vk|
        vk.vts.select { |vt|
          true # vt.tarmod != "B"
        }.to_a
      }.flatten
      @vts.each_with_index {|vt, i| vt.nr_im_vertrag = i+1}
    end
    @vts
  end

end

class DbfDat::Vt
  attr_writer :nr_im_vertrag
  def nr_im_vertrag
    if true # tarmod != "B" then
      vd.normale_vts
      @nr_im_vertrag
    end
  end
end



end # if not defined? Wandler_Hr5

if __FILE__ == $0 then
  durchlaufe_unittests($0)
end
