
if not defined?(Ort_Haupt) then

require 'win32ole'
# Auskommentiert, um Zeit (1 Sek) beim Starten zu sparen,
# die Bibliotheken werden dann bei Bedarf nachgeladen:
#    require 'models/system_fkt/dbase_zugriff'
#    require 'models/system_fkt/excel_zugriff'

require 'schmiedebasis'
require 'controllers/gui_apollo/gui_ap'
require 'controllers/prozesse/dienstherr'
#require 'models/orte'
require 'models/ortpool'
require 'models/system_fkt/speicherung'

#require 'guikonfig_ap'


class PfadExistiertNicht < RuntimeError
end

class Ort_Haupt < Ort_Gui
  def controller
    $ats
  end

  def self.neu(controller, vater, pfad_zum_neuen_ort)
    neu_pfad = pfad_zum_neuen_ort.gsub("//","/").chomp("/")
    bezeichnung = if vater
      File.basename(neu_pfad)
    else
      neu_pfad
    end
    ort_klasse = if bezeichnung =~ /<([^<>]*)>/
      case $1
      when /Excel-Dateien/i
        OrtExlDat
      when /\.xls$/i
        OrtExList
      when /gftest32\.exe/i
        OrtGftest
      else
        OrtDbase
      end
    else
#      File.directory?(neu_pfad)
#      bezeichnung.sub!(/^#{vater.ordner_pfad}/, '') if vater
      sub_pfade = Dir[neu_pfad+"/*"] # #*# ist nicht DRY mit kanonische subpfade
      hat_unterverzeichnisse =
        sub_pfade.find do |s| 
          (File.directory?(s) and s !~ /\/Referenzdaten$/i) or s =~ /\$\.xls$/i
        end
      hat_excel_dateien =
        sub_pfade.find { |s| s =~ /[^$]\.xls$/i }
      if hat_unterverzeichnisse then
        OrtOrdner
      else
        if hat_excel_dateien then
          OrtExlOrd
        else
          hat_dbase_dateien = sub_pfade.find { |s| s =~ /\.dbf$/i }
          hat_dbase_dateien ?
            OrtOrdner :
            OrtLeer
        end
      end
    end
    begin
      ort_klasse.new(controller, vater, bezeichnung)
    rescue PfadExistiertNicht
      nil
    end
  end

  def initialize(controller, vater, bezeichnung)
    @@controller = @controller = controller
    if not controller.respond_to?(:form) then
      trc_fehler :unzul_controller, [controller, vater, bezeichnung]
    end
    @@places = controller.form.places
    super(vater, bezeichnung)
    @gesamtanz_vertraege = 0
    @iterator_eintrag = nil

    if not File.exist?(kanonischer_pfad) then
      trc_info "nicht exist:", kanonischer_pfad
      trc_caller "aufr Ort_Haupt#ini", 7
      raise PfadExistiertNicht
    end

    if vater then
      controller.form.places.add_place(vater, self)
    else
      controller.form.places.add_mainplace(self)
    end
  end


  def spalten_defs
    [
      {:name=>:werte,     :text=>"Werte",    :align => Phi::TA_RIGHT_JUSTIFY},
      {:name=>:naja,      :text=>"Ung.",     :align => Phi::TA_RIGHT_JUSTIFY},
      {:name=>:schlechte, :text=>"Fehl.",    :align => Phi::TA_RIGHT_JUSTIFY},
      {:name=>:zeit,      :text=>"VglZeit",  :width=>60, :align => Phi::TA_CENTER},
      {:name=>:bemerkung, :text=>"Bemerkung",:width=>300}
    ]
  end

  def komplett_aufloesen
    super
    $ortpool.entferne(self.anzeige_pfad)
    @controller = nil
  end

  def komplett_aufloesen_sichtbar
    vater = @vater
    komplett_aufloesen
    vater.on_select if vater
  end


  def loesche_unterorte
    super
    self.node.delete_children if self.node
  end

public
  def neu_einlesen
    trc_temp :OrtRefresh_self_amAnfang, self
    snode = self.node
    trc_temp :selfnode_vorLoeschen, snode
    loesche_unterorte
    trc_info :selfnode_nachLoeschen, self.node
    self.generiere_unterorte
    trc_temp :selfnode_nachGenUnterorte, self.node
    snode.expand(false) # Wenigstens etwas, weil die Expandiertheit ja verloren geht.
    aktueller_ort.on_select if aktueller_ort
    trc_temp :verlasse_neueinlesen
  end

  def speicher_pfad
    anzeige_pfad
  end

  def on_select
    trc_info :ortselect_anfang_anzpfad=, anzeige_pfad
    #trc_temp caller[0..5].join("\n").to_s
    if not $orte_anzeigbar
      trc_info "Orte noch nicht anzeigbar."
      return
    end

    super

    if @iterator_eintrag
      sichtbaren_eintrag_aktualisieren(@iterator_eintrag, nil)
    end
    form.spez_menu = place_spezifisches_menu
    form.menu_wurzelort_entfernen.enabled = !(self.vater)
    form.menu_excel_vergleich_rekursiv.enabled = !(OrtDbase === self)
    form.anyfeld_panel.visible = !form.menu_excel_vergleich_rekursiv.enabled
    trc_info :ortselect_fertig
  end

  def ausfuehren
    aktivieren
  end

  attr_reader   :anzahl_vertraege_schon_ermittelt
  attr_reader   :aktualitaet_der_gesamtanz_vertraege
  attr_accessor :gesamtanz_vertraege


  def ermittle_anzahl_vertraege(zeitgrenze)
    #trc_temp :einsprung_self=, self
    return false if Time.now > zeitgrenze
    return true  if @anzahl_vertraege_schon_ermittelt
    yield_to_events
    trc_temp :beginne_self=, self
    if unterorte_erlaubt? then
      alle_unterorte_ok = true
      @gesamtanz_vertraege = untergeordnete.inject(0) do |anzahl, unterort|
        unterort.ermittle_anzahl_vertraege(zeitgrenze)
        anzahl += unterort.gesamtanz_vertraege if unterort.gesamtanz_vertraege
        alle_unterorte_ok = false unless unterort.anzahl_vertraege_schon_ermittelt
        #trc_temp :summe_unterort_anz, anzahl
        anzahl
      end
      @anzahl_vertraege_schon_ermittelt = alle_unterorte_ok
      setze_aktuellstempel_fuer_anzahl_vertraege
    else
      #trc_temp :direkt
      @gesamtanz_vertraege = direkt_enthaltene_vertraege
    end
    trc_temp :ergebnis_self_anz, [self, @gesamtanz_vertraege]
    true
  end

protected
  def setze_aktuellstempel_fuer_anzahl_vertraege
    @aktualitaet_der_gesamtanz_vertraege = Time.now
  end

  def direkt_enthaltene_vertraege=(anzahl)
    setze_aktuellstempel_fuer_anzahl_vertraege
    @gesamtanz_vertraege = @direkt_enthaltene_vertraege = anzahl
    trc_temp :direkt=, @direkt_enthaltene_vertraege
    @anzahl_vertraege_schon_ermittelt = true
    anzahl
  end


public
  def place_spezifisches_menu
    trc_info :default_spez_menu_abruf#, @@spezifisches_menu
    nil
  end

  def self.vliste_menu
    allg_vliste_menu +    [

     ]
  end

  def self.aktueller_ort
    @@places.aktiver_ort
  end

  def aktueller_ort
    self.class.aktueller_ort
    #form.places.aktiver_ort
  end


public
  def konsole
    controller.konsole
  end

  def persist_proxy
    @persist_proxy ||= OrtPersistenz.fuer_pfad(speicher_pfad)
  end

  def eintrag_persist(eintragsname)
    persist_proxy.inhalt_fuer_eintrag(eintragsname)
  end

  def laden
    persist_proxy
  end

  def speichern
    persist_proxy.speichern
  end

  def setze_kurzinfo(eintragsname, zeile)
    eintrag_persist(eintragsname).kurzinfo = zeile
    speichern
  end

  def kurzinfo(eintragsname)
    eintrag_persist(eintragsname).kurzinfo
  end

  def ziel_ort=(z)
    eintrag_persist(".").zielort = z.anzeige_pfad
    speichern
    @ziel_ort = z
  end

  def ziel_ort
#    if @ziel_ort.nil? then
      @ziel_ort = $ortpool.pfad_zu_ort(eintrag_persist(".").zielort, :selbst_als_wurzel_wenn_keine_vorfahren)
      @ziel_ort ||= false # wird also nie wieder nil sein
 #   end
  #  @ziel_ort
  end

  def quell_ort=(q)
    eintrag_persist(".").quellort = q.anzeige_pfad
    speichern
    @quell_ort = q
  end

  def quell_ort
#    if @quell_ort.nil? then
      @quell_ort = $ortpool.pfad_zu_ort(eintrag_persist(".").quellort, :selbst_als_wurzel_wenn_keine_vorfahren)
      @quell_ort ||= false # wird also nie wieder nil sein
 #   end
  #  @quell_ort
  end


  def neues_vglerg(eintragsname, vglerg)
    if vglerg.respond_to? :zeit then
      persist_proxy.inhalt_fuer_eintrag(eintragsname).vglerg_neu_werte(vglerg)
      speichern
    else
      trc_hinweis :kein_vglerg, vglerg
    end
  end

  # erg:: nil bedeutet, dass die Berechnung gerade anfï¿½ngt
  # (dass also noch kein Ergebnis ermittelt wurde).
  def sichtbaren_eintrag_aktualisieren(eintrag, erg)
    eintr_name = eintrag.respond_to?(name) ? eintrag.name : eintrag
#    trace :sichtbaresaktualisieren_ort, self
  #  trace :sichtbaresaktualisieren_akt, aktueller_ort
    trc_info :sichtbaresaktualisieren_name, eintr_name

    if self.equal?(aktueller_ort)
      #@gerade_aktueller_eintrag = nil
      ##*# Phi-Funktion nicht so roh aufrufen:
      listitem = form.vliste.find_caption(0, File.basename(eintr_name), false, true, true)
      if listitem
        trc_temp :listitem_gefunden_erg=, erg
        if not erg then
          #@gerade_aktueller_eintrag = eintr_name
          #@altes_icon_von_eintrag = listitem.state_index # #+# funktioniert nicht bei doppeltem Aufruf.   #?*# wrappen?
        end
        visu = {:icon=>erg_zu_icon(erg)} #, @altes_icon_von_eintrag)}
        case erg
        when VergleichsErgebnis
          visu.update(:zeit=>Time.now.strftime("%H:%M")).
               update(zusatz_listeninhalt_fuer_erg(erg))
        end
        listitem.setze_inhalt(visu)
      else
        if form.vliste.items.count > 0 then
          trc_hinweis :listitem_nichtgefunden_vliste0, form.vliste.items[0]
        else
          trc_hinweis :listitem_nichtgefunden_vlistecount, form.vliste.items.count
        end
      end
    else
      trc_temp :eintrag_nicht_im_aktuellen_ort_aktort=, aktueller_ort
    end
    #form.als_vorderstes_fenster
    yield_to_events
  end

  alias :eintrag_aktualisieren :sichtbaren_eintrag_aktualisieren


  def erg_zu_icon(erg, altes_icon=IMGIDX_LEER)
#    trc_temp :erg, erg
    case erg
    when nil
      IMGIDX_AKTIV
    when VergleichsErgebnis
      AMPEL_ICONS[erg.ampelwert]
    when 1
      altes_icon
    when 0
      IMGIDX_ROT
    when -9999 .. -1
      IMGIDX_LILA
    when :fehler, false
      IMGIDX_LILA
    else
      trc_info "unbekanntes Ergebnis beim eintrag-aktualisieren", erg
      IMGIDX_LILA
    end
  end

  def eigene_ansicht_aktualisieren(erg)
    if not erg then
      self.aktivieren if KONFIG.opts[:OrtsAnzeigeSynchronisieren]
      form.on_idle.call
    end
    node.state_index = erg_zu_icon(erg, IMGIDX_UNDEF)
    if self.vater then
      vater.sichtbaren_eintrag_aktualisieren(self.name, erg)
    end
  end

  def durchlauf_ende(kumuliertes)
    eigene_ansicht_aktualisieren(kumuliertes)
    trc_temp :kumul, kumuliertes
  end

  def zusatz_listeninhalt_fuer_erg(erg)
    #trc_temp :vgl, erg
    ges = fehl = ""
    case erg
    when VergleichsErgebnis then
      ges  = erg.summe
      fehl = erg[:eaVollDaneben, :eaXception]
      naja = erg[:eaNochOK, :eaUngenau]
      {:werte=>ges, :naja=>naja, :schlechte=>fehl}
    else
      {}
    end
  end

  def zusatz_inhalt(range = 0..-1)
    untergeordnete[range].map do |eintrag|
      eintrag_name = if eintrag.respond_to?(:name)
        eintrag.name
      else
        File.basename(eintrag.to_s)
      end
#      trc_temp :zusatz_inhalt_eintragname, eintrag_name
      zusatz_inhalt_fuer_item(eintrag_name)
    end
  end

  def zusatz_inhalt_fuer_item(dateiname)

    eintrag_ps = persist_proxy.inhalt_fuer_eintrag(dateiname)
#    trace :vglebnisse, eintrag_ps

    erg = if false #@testzeit_combx.selectedString > 0 then
      trc_temp :combx, @testzeit_combx.getTextOf(@testzeit_combx.selectedString)
      zeit = @testzeit_combx.getTextOf(@testzeit_combx.selectedString)
      eintrag_ps[zeit]
    else
      eintrag_ps.vglerg_liste_werte.sort_by {|ve| ve.zeit}.last
    end
    aktuelles_jahr = Time.now.year
    zeit = erg ? erg.zeit.to_s : ""
    alles, jahr_str, monat_str, tag_str, uhrzeit_str =
      zeit.match(/^(\d{2,4})-(\d{1,2})-(\d{1,2})_(\d?\d:?\d\d)$/).to_a
    #.sub(/#{aktuelles_jahr}[-]?/,"")
    #trc_temp :zusatzinhalt_zeit, [zeit, alles, jahr_str, monat_str, tag_str, uhrzeit_str]
    if zeit != "" then
      zeit = Time.local(aktuelles_jahr,monat_str.to_i).strftime("%b") + "-" + tag_str # Abgek. Monatsnamen
      if jahr_str.to_i != aktuelles_jahr then
        zeit = jahr_str + zeit.sub(/-/,"")
      else
        if monat_str.to_i == Time.now.month and tag_str.to_i == Time.now.day then
        #if zeit.sub!(/^#{Time.now.strftime("%m-%d")}[_]?/,"")
          zeit = uhrzeit_str[0,2] + ":" + uhrzeit_str[-2,2]
        end
      end
    end
    erg ||= 1 # für die Icons wï¿½rde nil bedeuten, dass das Ergebnis gerade berechnet wird
    icon = if @iterator_eintrag == dateiname then
      IMGIDX_AKTIV
    else
      erg_zu_icon(erg)
    end
    zusatz_listeninhalt_fuer_erg(erg).update(
      :icon => icon,
      :zeit => zeit,
      :bemerkung => eintrag_ps.kurzinfo
    )
  end

  def starte_dienst(dienstherr_klasse, *andere_args, &abschluss_proc)
    begin
      dienstherr = dienstherr_klasse.new(self.controller, self, *andere_args, &abschluss_proc)
    rescue
      dienst_name = begin
        dienstherr_klasse.name.sub(/herr/i,"")
      rescue
        "Dienst (#{dienstherr_klasse.inspect})"
      end
      trc_fehler "Dienst-Ort", self
      trc_fehler "Dienst-Args", andere_args
      controller.melde_aktuellen_error "\n#{dienst_name} nicht gestartet"
      messageBox("Dienst nicht gestartet\n\n" + $!.to_s)
    end
  end

  def standard_ziel_pfad(ziel_ort)
    ziel_ort.ordner_pfad + File.basename(self.anzeige_pfad)
  end

  def bewege_nach(ziel_pfad, option=nil)
    if true then
      ziel_ort = $ortpool.pfad_zu_ort(ziel_pfad)
      trc_info :vor_refresh, ziel_ort
      ziel_ort.neu_einlesen
    end
  end

  def physisch_loeschen
    antwort = Phi::message_dlg(<<EOT,
!! Achtung !!

Der ausgewï¿½hlte Ort (#{self.anzeige_pfad})
wird jetzt komplett mit Inhalt von der Festplatte gelï¿½scht.

Fortfahren?
EOT
               Phi::MT_CONFIRMATION,
               [Phi::MB_YES, Phi::MB_NO], 0 )
   trc_info :ortphysloesch_antwort1, antwort

   return false unless antwort == Phi::MR_YES

    pos = Phi::get_cursor_pos #form.screen_to_client
    trc_temp :pos, pos
    antwort = Phi::message_dlg(<<EOT,
!! Es geht um echtes Lï¿½schen von der Festplatte !!

Soll der Ort #{self.anzeige_pfad}
wirklich komplett mit Inhalt von der Festplatte geï¿½scht werden?
Eventuelle Unterordner werden damit auch gnadenlos vernichtet.

Wirklich fortfahren?
EOT
               Phi::MT_CONFIRMATION,
               [Phi::MB_OK, Phi::MB_CANCEL],
               0, #100, 100)
               pos.x/4, pos.y/4)
    trc_info :ortphysloesch_antwort2, antwort
    antwort == Phi::MR_OK
  end

  def generiere_unterorte(auswahl_pfad = nil)
    #trc_temp :guo_start_auswpfad_kanonpfad, [auswahl_pfad, self.anzeige_pfad]
    kan_subpfade = kanonische_subpfade
    unless kan_subpfade.empty? or
           kan_subpfade.size == 1 and kan_subpfade[0] =~ /\/\*\.xls$/i
      kan_subpfade.each do |subpfad|
        new_place = generiere_oder_finde_unterort(subpfad)
        #trc_temp :new_place, new_place
        next unless new_place # #*# sollte aber nur selten vorkommen!
        #next unless new_place
        sp = subpfad.chomp("/")
        #trc_temp :ue_subpfad, [new_place.unterorte_erlaubt?, sp]
        next if auswahl_pfad and not sp == auswahl_pfad[0, sp.size]
        new_place.generiere_unterorte(auswahl_pfad) if new_place.unterorte_erlaubt?
      end
    end
    self.gesamtanz_vertraege
  end

protected
  def self.spez_menu_cache
    trc_temp :spez_menu, [self, @spez_menu]
    @spez_menu ||= yield
  end

  def statuszeile
  end

public
#private #disabled wg tests 
  
  def generiere_oder_finde_unterort(unterpfad)
    return $ortpool.pfad_zu_ort(unterpfad, self)

=begin
    erg = direkter_unterort_per_pfad(unterpfad)
    return erg if erg
    neuer_ort = Ort_Haupt.neu(self.controller, self, unterpfad)
    if neuer_ort.nil? then
      trc_fehler "!!! Ort.new hat nil zurï¿½ckgegeben", unterpfad
    else
      self.gesamtanz_vertraege += neuer_ort.gesamtanz_vertraege
    end
    neuer_ort
=end

  end

  def kanonische_subpfade
    Dir[(anzeige_pfad+"*")].sort.map { |subpfad|
      if subpfad =~ /gftest32\.exe$/i then
        AnatolsTestprog.hinzu(subpfad)
      end
      if File.directory?(subpfad)
        next if subpfad =~ /\/Referenzdaten$/i
        subpfad
      else
        dateiname_zu_anzeigepfad(subpfad)
      end
    }.compact.uniq
  end

end


#**************************************************
#



end # if not defined?(Ort_Haupt)

if __FILE__ == $0
  durchlaufe_unittests($0)
end
