
if not defined?(OrtOrdner)

require 'schmiedebasis'
require 'models/ortpool'
require 'models/ort_haupt'

# auch diese Dateien werden hier includiert, damit orte vollst.
# #*# Aufrï¿½umen!!!!!!!!!!!!!!!!!!!!!!!!!!
require 'models/ort_excel' 
require 'models/ort_exlist' 


#**************************************************
#


class OrtDbase < Ort_Haupt

  def profil_name
    name. sub(/^</,""). sub(/>$/,"")
  end


  def ordner_pfad
    File.dirname(anzeige_pfad)+"/"  # also ohne den Profilnamen
  end

  def kanonischer_pfad
    ordner_pfad + "vdB#{profil_name}.dbf"
  end
  def generiere_unterorte(auswahl_pfad = nil)
    gesamtanz_vertraege
  end

  def spalten_defs
    [
      {:text => "vsnr",   :width => 100},
      {:name=>:ref,       :text => "Ref", :width => 35, :align => Phi::TA_CENTER},
      {:text => "beiart", :align => Phi::TA_RIGHT_JUSTIFY},
      {:name=>:anyfeld,   :text => form.anyfeld.text,   :width => 50}
    ] + super + [
    ]
  end
  def self.image_index
    IMGIDX_ORDNERDBASE
  end
  def untergeordnete
    require 'models/system_fkt/dbase_zugriff'
    DbfDat::oeffnen( kanonischer_pfad, :b )
    vdname  = File.basename(kanonischer_pfad).sub(/\.dbf$/, '')
    v_liste = DbfDat::schnell_lesen(vdname, "vsnr") #.fetch_all #T_VD LEFT OUTER JOIN #{vpname} T_VP ON T_VD.vsnr = T_VP.vsnr
    v_liste.map { |ele| anzeige_pfad + ele["vsnr"]}
  end

  def direkt_enthaltene_vertraege
    if not @direkt_enthaltene_vertraege then
      require 'models/system_fkt/dbase_zugriff'

      DbfDat::oeffnen( kanonischer_pfad, :b )
      trc_temp :kanonischer_pfad, kanonischer_pfad
      vdname  = kanonischer_pfad.sub(/\.dbf$/, '').gsub("/", "\\")
      begin
        self.direkt_enthaltene_vertraege = DbfDat::schnell_anzahl(kanonischer_pfad)
        trc_info :Table_count, @direkt_enthaltene_vertraege
        #trc_temp caller[0..5].join("\n")
      rescue
        trc_aktuellen_error "DB-Fehler", 8
        self.direkt_enthaltene_vertraege = 0
      end
    end
    @direkt_enthaltene_vertraege
  end


  def inhalt
    require 'models/system_fkt/dbase_zugriff'

    dbfpfad = ordner_pfad

    vdname  = File.basename(kanonischer_pfad).sub(/\.dbf$/, '')  # "vdB#{profil_name}.dbf"
    vpname = vdname.sub(/^vd/i, "vp")

    DbfDat::oeffnen( dbfpfad + "/" + vdname, :b )
    trc_temp :dbfliste_fuer, dbfpfad
    begin
      #v_liste = DbfDat::sql_select("SELECT VSNR,BEIART FROM #{vdname}") #.fetch_all #T_VD LEFT OUTER JOIN #{vpname} T_VP ON T_VD.vsnr = T_VP.vsnr
      v_liste = DbfDat::schnell_lesen(vdname, "vsnr", "beiart") #("SELECT VSNR,BEIART FROM #{vdname}") #.fetch_all #T_VD LEFT OUTER JOIN #{vpname} T_VP ON T_VD.vsnr = T_VP.vsnr
      trc_temp :vliste, v_liste
      erg = v_liste.map { |ele| [IMGIDX_DBASE, ele["vsnr"], "", ele["beiart"], "", ""] }
      self.direkt_enthaltene_vertraege = erg.size
      erg
    rescue ActiveRecord::StatementInvalid
      trc_fehler "dbf-liste #{vdname}", $!
      [[IMGIDX_DBASE, 'Datei nicht gefunden', '', '', '']]
    end
  end


  def zusatz_inhalt(range = 0..-1) #_fuer_item(vsnr)
    vd_ganzer_name  = kanonischer_pfad.sub(/\.dbf$/, '')
    vd_basename = File.basename(vd_ganzer_name)
    DbfDat::oeffnen( vd_ganzer_name, :S )
    #vd = nil
    begin
      # Achtung, Gefahr von Speicherloch!
      @vds ||= DbfDat::Vd.find(:all) #(vsnr) #  DbfDat::sql_select("SELECT VSNR FROM #{vd_basename} WHERE VSNR='#{vsnr}'") #.fetch_all #T_VD LEFT OUTER JOIN #{vpname} T_VP ON T_VD.vsnr = T_VP.vsnr
      #trace :vd_aus_refdaten, vd
    rescue ActiveRecord::StatementInvalid, NoMethodError,
           RuntimeError # das letzte nur, weil ich im BDE-Adapter noch keine eigenen Exception-Klassen habe.
      trc_hinweis :dbfliste_error, $!
      return nil
    end

    trc_temp "vds.size",  @vds.size

    untergeordnete[range].map do |pfad|
      vsnr = File.basename(pfad)
      vd = @vds.find {|vd| vd.vsnr == vsnr}
      ref = vd ? "S" : ""
      zusatz_inhalt_fuer_item(vsnr).update({:ref=>ref})
    end
  end

  FELD_SELEKTOREN = {
      "vd" => proc { |vd, f| f.call(vd)},
      "vp" => proc { |vd, f| f.call(vd.vp)},
      "st" => proc { |vd, f| f.call(vd.st)},
      "vk" => proc { |vd, f| vd.vks.map { |vk| f.call(vk) }},
      "vt" => proc { |vd, f| vd.vks.map { |vk| vk.vts.map { |vt| f.call(vt) }}}
    }

  def zusatz_inhalt_langsam(vsnr)
    anyfeld_def = form.anyfeld.text
    werte =
      if anyfeld_def > ""
        alles, ziel, rest_def = anyfeld_def.match(/^(v.)\.(.+)/).to_a
        begin
          @vds ||= Vd.find(:all)
          vd = @vds.find {|vd| vd.vsnr == vsnr}
          if vd then
            trc_temp :alles___rest_def, [alles, ziel, rest_def]
            if ziel && rest_def
              trc_temp :rest_def, rest_def
              FELD_SELEKTOREN[ziel].call(vd, proc { |obj|
                #trace :obj, obj
                eval("#{ziel}=obj;#{ziel}.#{rest_def}")
              })
            else
              eval(anyfeld_def)
            end

          end
        rescue Exception
          trc_hinweis "anyfeld #{anyfeld_def}", $!
          "--"
        end
      end
    werte ||= ""
    werte = [werte].flatten.map do |wert|
      trc_temp :wert, wert
      if wert.nil?
        "nil"
      elsif wert.respond_to?(:to_i)
        ((wert==wert.to_i) ? wert.to_i : wert ).to_s
      elsif wert == ""
        "\"\""
      elsif wert.is_a?(String)
        wert.gsub(" ","_")
      else
        wert.to_s
      end
    end
    anyfeld = werte.join("; ") # if anyfeld.is_a?(Array)

    {:anyfeld=>anyfeld}
  end

  def on_vertr_dbl_click(item)
    trc_fehler :dbf_vertragsanzeige_deaktiviert, self
    return if self
    require 'guivertr_ap'
    voller_dateiname = anzeige_pfad + '/' + item.caption
    Vertrag.anzeigen( voller_dateiname )
  end

  def place_spezifisches_menu
    trc_temp :dbase_spez_menu_abruf
    self.class.spez_menu_cache {
        trc_temp :dbase_spez_menu_create
        form.make_main_menu(:dbase_menu,
          [{
            :text=>'dBase',
            :name=>:menu_dbase,
            :menu=>[{
                :text=>'Excel-Testfälle in einzelnen Dateien erzeugen (NG1)',
                :name=>:menu_dbase_excel_ng1,
                :proc=>proc{ erzeuge_excel(:NG1) }
              },{
                :text=>'Excel-Testfälle in einzelnen Dateien erzeugen (HR5)',
                :name=>:menu_dbase_excel_hr5,
                :proc=>proc{ erzeuge_excel(:HR5) }
              },{
                :text=>'Excel-Testfälle in Serien-Datei erzeugen (HR6)',
                :name=>:menu_dbase_excel_hr6,
                :proc=>proc{ erzeuge_excel(:HR6) }
              },{
                :text=>'Excel-fälle erzeugen und vergleichen (NG1)',
                :proc=>method(:erzeuge_excel_und_vergleiche)
              }]
          }]
        )
    }
  end

  def physisch_loeschen
    return unless super
    #require 'ftools'
    #ordner = ziel_place.realpath + "/" + File.basename(self.scheinbarer_pfad)
    require 'models/system_fkt/dbase_zugriff'
    DbfDat::schliessen
    self.parent_ort.aktivieren
    DbfDat::VO_KLASSEN.each do |vklasse|
      id = vklasse.name.downcase
      %w[B S].each do |s_oder_b|
        zielname = self.ordner_pfad + id + s_oder_b + self.profil_name+".dbf"
        trc_info "loesche Datei", zielname
        begin File.delete zielname rescue nil end
      end
    end
    self.komplett_aufloesen_sichtbar
  end

  def bewege_nach(ziel_pfad, option=nil)
    require 'ftools'
    require 'models/system_fkt/dbase_zugriff'
    konsole.meldung "\n"

    arten = case option
      when :ref  then %w[S]
      when :work then %w[B]
      else            %w[B S]
    end
    DbfDat::VO_KLASSEN.each do |vklasse|
      id = vklasse.name.split('::').last.downcase
      arten.each do |s_oder_b|
        quellname = self.kanonischer_pfad.sub(/\/vd[sb]/i,"/"+id+s_oder_b)
        zielname = File.dirname(ziel_pfad) + "/" +
                  File.basename(ziel_pfad).sub(/^<?/,id+s_oder_b).sub(/>?$/,"") + ".dbf"
        trc_info "q_z", [quellname, zielname]
        begin
          File.copy quellname, zielname
        rescue
          trc_info "nicht kopiert", $!.to_s
          konsole.meldung "Ausnahme: "+File.basename(quellname) + " nicht kopiert." #(#{$!.to_s})"
        end
      end
    end
    begin
      File.copy persist_proxy.dateiname, File.dirname(ziel_pfad)+"/"+File.basename(persist_proxy.dateiname)
    rescue
      trc_info "persist nicht kopiert", $!.to_s
    end

    konsole.meldung "\nProfil nach #{ziel_pfad} kopiert."
    #  "+self.anzeige_pfad+ "
    super
  end
    
  # art ist :HR6 oder :NG1
  def erzeuge_excel(art)  
    trc_temp  :menu_dbase_excel_on_click_art=, art
    
    return if form.vliste.items.count == 0
    
    if form.vliste.sel_count == 0
      messageBox "Es waren keine Vertrï¿½ge ausgewï¿½hlt,\nalle Vertrï¿½ge werden verwendet."
      form.vliste.select_all
    end

    if self != aktueller_ort
      return aktueller_ort.erzeuge_excel(art)
    end

    exstar_vorlage_auf_neusten_stand
    require 'models/system_fkt/excel_zugriff'

    zielpfad_sym = "Dbf2ExlAusgabePfad_#{art}".to_sym
    ziel_pfad = KONFIG.opts[zielpfad_sym]
    loop do
      
      dlg = EinstellungenDlg.new("Dbf_zu_Excel_#{art}".to_sym) do |was, neu_wert|
        trc_temp :lbl_update, neu_wert
        ziel_pfad = neu_wert if was == zielpfad_sym
        "Erzeuge Exstar-Testfälle aus Mathstar-Profil #{self.name}\n" +
        "in " + if art.to_s=="NG1" then 
          ziel_pfad + "/"+profil_name 
        else 
          ziel_pfad.sub(/([^$])(\.xls)$/, '\1$\2')
        end
      end
      erg = dlg.show_modal
      trc_info :mr_result, erg
      return if erg == Phi::MR_CANCEL

      next if not File.exist?(KONFIG.opts["ExstarVorl_#{art}".to_sym])
      next if KONFIG.opts[:VorgBibliothekNutzen]  and
                (KONFIG.opts[:VorgegebeneBibliothek] == "" or
                    not ExcelZugriff.existiert_vorgegebene_bibbliothek?)

      break
    end 

    trc_temp :erz_excel__ort=, self
 #   default_ordner = (KONFIG.opts[:Dbf2ExlAusgabePfad]||WORK_DIRNAME) + "/" + profil_name
#    ordner = default_ordner #form.ordner_abfragen("Ordner für die Excel-Dateien:", "", default_ordner)
    ziel_pfad = KONFIG.opts[zielpfad_sym]
    case art.to_sym
    when [:NG1, :HR5]
      ordner = ziel_pfad.chomp("/") + "/" + profil_name
      trc_info :ordner, ordner
      return unless ordner
      ref_ordner = ordner + "/Referenzdaten"
      Phi::force_dirs(ref_ordner)
      trc_info :ref_ordner, ref_ordner
  
      zielort = $ortpool.pfad_zu_ort(ordner, :selbst_als_wurzel_wenn_keine_vorfahren)
    when :HR6
      if File.directory?(ziel_pfad) then
        ziel_pfad += "/" + profil_name + "$.xls"
      elsif ziel_pfad !~ /\.xls$/i then
        ziel_pfad += "$.xls" 
      else 
        ziel_pfad.sub!(/([^$])(\.xls)$/, '\1$\2')
      end
      Phi::force_dirs(File.dirname(ziel_pfad))
      schon_vorhanden = File.exist?(ziel_pfad)
      File.open(ziel_pfad,"w") {} unless schon_vorhanden 
      zielort = $ortpool.dateipfad_zu_ort(ziel_pfad, :selbst_als_wurzel_wenn_keine_vorfahren)
      File.delete ziel_pfad unless schon_vorhanden
    end
      
    vsnr_liste = form.vliste.items.map { |item|
        item.caption if item.selected
      }.compact

    starte_dienst(Dienstherr_Dbf2Exl, zielort, vsnr_liste, art) do
      trc_info :beginne_weidereinlesen, ordner
      zielort.neu_einlesen
=begin #alter code      
      ziel_ortkonfig = controller.sortierte_ortskonfig.find do |(name, inhalt)|
        trc_temp :inhalt, inhalt
        start_ort = inhalt[:place]
        trc_temp :start_ort_pfad, start_ort
        start_ort.finde_ort_per_pfad(basis_ordner) if start_ort
      end
      trc_temp :ziel_ort_gefunden, ziel_ortkonfig
      if ziel_ortkonfig then
        name, inhalt = ziel_ortkonfig
        ziel_ort = inhalt[:place]
        ziel_ort.neu_einlesen
      end
=end
    end
  end

  def erzeuge_excel_und_vergleiche(sender)

    trc_temp  :menu_dev_on_click
    if form.vliste.sel_count == 0
      messageBox "Es sind keine Vertrï¿½ge ausgewï¿½hlt\n zum erstellen von Excel-Daten."
      return
    end

    if self != aktueller_ort
      return aktueller_ort.erzeuge_excel_und_vergleiche(sender)
    end

    exstar_vorlage_auf_neusten_stand
    require 'models/system_fkt/excel_zugriff'

    art = :NG1
    zielpfad_sym = "Dbf2ExlAusgabePfad_#{art}".to_sym
    ziel_pfad = KONFIG.opts[zielpfad_sym]
    loop do
      dlg = EinstellungenDlg.new("Dbf_zu_Excel_#{art}".to_sym) do |was, neu_wert|
        trc_temp :lbl_update, neu_wert
        ziel_pfad = neu_wert if was == zielpfad_sym
        "Erzeuge Exstar-Testfälle aus Mathstar-Profil #{self.name}\n" +
        "und vergleiche diese.\n\n" +
        "Ziel-Ordner:\n" +
        ziel_pfad + if art.to_s=="NG1" then "/"+profil_name else "" end
      end
  
      erg = dlg.show_modal
      trc_info :mr_result, erg
      return if erg == Phi::MR_CANCEL

      next if not File.exist?(KONFIG.opts["ExstarVorl_#{art}".to_sym])
      next if KONFIG.opts[:VorgBibliothekNutzen]  and
                (KONFIG.opts[:VorgegebeneBibliothek] == "" or
                    not ExcelZugriff.existiert_vorgegebene_bibbliothek?)

      break
    end

    trc_temp :erz_excel__ort=, self
 #   default_ordner = (KONFIG.opts[:Dbf2ExlAusgabePfad]||WORK_DIRNAME) + "/" + profil_name
#    ordner = default_ordner #form.ordner_abfragen("Ordner für die Excel-Dateien:", "", default_ordner)
    basis_ordner = KONFIG.opts[zielpfad_sym]
    ordner = basis_ordner.chomp("/") + "/" + profil_name
    trc_info :ordner, ordner
    return unless ordner

    ref_ordner = ordner + "/Referenzdaten"
    Phi::force_dirs(ref_ordner)
    trc_info :ref_ordner, ref_ordner

    vsnr_liste = form.vliste.items.map { |item|
        item.caption if item.selected
      }.compact

    ziel_ort = $ortpool.pfad_zu_ort(ordner, :selbst_als_wurzel_wenn_keine_vorfahren)

    starte_dienst(Dienstherr_DbfExlVgl, ziel_ort, vsnr_liste) do
      trc_info :beginne_weidereinlesen, ordner
      ziel_ortkonfig = controller.sortierte_ortskonfig.find do |(name, inhalt)|
        trc_temp :inhalt, inhalt
        start_ort = inhalt[:place]
        trc_temp :start_ort_pfad, start_ort
        start_ort.finde_ort_per_pfad(basis_ordner) if start_ort
      end
      trc_temp :ziel_ort_gefunden, ziel_ortkonfig
      if ziel_ortkonfig then
        name, inhalt = ziel_ortkonfig
        ziel_ort = inhalt[:place]
        ziel_ort.neu_einlesen
      end
    end
  end

  def vergleichen
    VergleichsErgebnis.new # d.h. nicht implementiert
  end

end



#**********************************************************

class OrtOrdner < Ort_Haupt
  def spalten_defs
    [
      {:text => "Ort", :width => 150},
      {:text => "Anz", :align => Phi::TA_RIGHT_JUSTIFY}
    ] +
      super
  end

  def self.image_index
    IMGIDX_ORDNER
  end

  def direkt_enthaltene_vertraege
    0
  end

  def inhalt
    subplaces.map {|place|
      anz_vertr = (place.gesamtanz_vertraege if
                      place.anzahl_vertraege_schon_ermittelt or
                      place.gesamtanz_vertraege > 0)
      [place.class.image_index,
       place.name,
       anz_vertr
      ]
    }
  end

  def self.vliste_menu
    [
      {:text => "&Lï¿½schen",
       :proc => proc {|ort| ort.physisch_loeschen}}
    ]
  end

  def untergeordnete
    subplaces
  end

  def unterorte_erlaubt?
    true
  end

  def on_vertr_dbl_click(item)
    neuer_ort = self.unterort_per_name(item.caption)
    neuer_ort.ausfuehren if neuer_ort
  end

  def ok_als_ziel?(quelle)
    #trc_temp :quelle, quelle
    not quelle.is_a? OrtGftest
  end

  def bewege_nach(zielname, option=nil) # ziel_place)
    #zielname = super
    quellname = self.ordner_pfad.sub(/\/$/,"")

    require 'ftools'

    return unless zielname
    trc_temp  "q_z", [quellname, zielname]
    quellname, zielname = [quellname, zielname].map do |name|
      name.gsub(/\//,"\\") #.gsub(/ /,"\\ ")
    end
    #nirwana = LOG_DIRNAME + "/xcopy_ausgabe.txt"
    befehlszeile = %Q(xcopy /E/I/Y "#{quellname}" "#{zielname}" >NUL)
    trc_info befehlszeile
    trc_hinweis :xcopy, system(befehlszeile)

    super #ziel_place.neu_einlesen
  end

  def physisch_loeschen
    return unless super
    Dir.rm_rf(self.ordner_pfad)
    self.komplett_aufloesen_sichtbar
  end

end

#**********
class AnatolsTestprog
  @@alle_testprogs = {}

  attr_reader :voller_progname

  def self.hinzu(voller_progname)
    @@alle_testprogs[voller_progname] || new(voller_progname)
  end

  def initialize(voller_progname)
    @voller_progname = voller_progname
    @@alle_testprogs[@voller_progname] = self
  end

  def <=>(anderes_testprog)
    self.voller_progname <=> anderes_testprog.voller_progname
  end

  def self.alle
    @@alle_testprogs.values
  end

  def self.zeige_menu
    trc_info :testprog_menu_start
    menu_spez = if alle.empty?
      [:text=>"Kein GfTest32.exe gefunden."]
    else
      alle.sort.map do |testprog|
        {:text=>testprog.voller_progname,
         :proc=>proc do
            IO::popen testprog.voller_progname
          end
        }
      end
    end
    MainForm::zeige_menu_an_mauspos(menu_spez)
  end
end


class OrtGftest < Ort_Haupt

  def aktivierbar?
    false
  end

  def on_dbl_click
    super
    ausfuehren
  end

  def ausfuehren
    IO::popen( trc_temp( :pfad, File.dirname(self.ordner_pfad) +  "/gftest32.exe"))
    trc_temp :gftest_gestartet
  end

  def self.image_index
    IMGIDX_GFTEST
  end

  def spalten_defs
    [
      {:text => "Pfad",   :width => 200}
    ] +
    super
  end

  def inhalt
    [[IMGIDX_DBASE, self.ordner_pfad, '']]
  end

  def direkt_enthaltene_vertraege
    0
  end
end


#**********************************************
#

class OrtLeer < Object
  self.public_methods.each do |methname|
    if methname !~ /(^__|method)/ then
      trc_temp :lehremove, methname
      undef_method(methname) rescue trc_temp :lehrefail, $!
    end
  end

  def method_missing(meth, *args, &blk)
    #trc_temp :lehr_mm, [meth, args]
    @real_ort.send(meth, *args, &blk)
  end

  attr_accessor :real_ort

  def initialize(*args)
    @real_ort = OrtOrdner.new(*args)
    proxy_ort = self

    ec = class << @real_ort; self; end
    ec.send(:define_method, :persist_proxy) do
        trc_temp :lehr_proxy, ordner_pfad
        # #*# ist nicht DRY mit kanonische subpfade und neu(...)
        hat_interessanten_inhalt = Dir[ordner_pfad+"*"].find do |s|
          trc_temp :lehr_inhalt, s
          s =~ /\.xls$/i or
          (File.directory?(s) and s !~ /\/Referenzdaten$/i )
        end

        if hat_interessanten_inhalt then
          self.node.delete
          trc_temp "ordner_pfad, vater", [ordner_pfad, vater]
          r = Ort_Haupt.neu(controller, vater, ordner_pfad)
          if r then
            proxy_ort.real_ort = r
            trc_temp :lehr_ersetzt, r
          end
          proxy_ort.real_ort.persist_proxy
    #      r.einlesen
     #     r.speichern
      #    trc_temp :leer_neu_gespeichert
        else
          super
        end
    end

  end

end

end # if not defined?(Ort_Ordner)

if __FILE__ == $0
  durchlaufe_unittests($0)
end
