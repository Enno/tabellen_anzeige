

if not defined?(OrtExcelBase)

require 'schmiedebasis'
require 'models/ortpool'
require 'models/ort_haupt'



#*************************************************************


class OrtExcelBase < Ort_Haupt

  def speicher_pfad # stimmt bei ExlDat mit anzeige_pfad überein, bei ExlOrd nicht!!
    ordner_pfad + "<Excel-Dateien>/"
  end

  def kanonischer_pfad
    ordner_pfad + File.basename(untergeordnete.first||"")
  end

  def generiere_unterorte(auswahl_pfad = nil)
    gesamtanz_vertraege
  end

  def spalten_defs
    [
      {:text => "Dateiname", :width => 150},
      {:name=>:kbytes,    :text => "Groesze", :width=>60, :align => Phi::TA_RIGHT_JUSTIFY},
      {:name=>:ref,       :text => "Ref", :width => 33, :align => Phi::TA_CENTER}
    ] +
      super
  end

  def self.image_index
    IMGIDX_ORDNEREXCEL
  end

  def pfad_dateimuster
    ordner_pfad + "*.xls"
  end

  def base_datei_regex
    @@base_datei_regex ||= /\/[^#{excel_zugriff::MOEGLICHE_REFPREFIXE[0,3].join("")}][^\/]*\.xls$/i
  end

  def untergeordnete
    #trc_temp :exl_untergeordnete_of, ordner_pfad
    dateiliste = Dir[self.pfad_dateimuster].select { |ganzername|
      ganzername =~ base_datei_regex
    }.sort
  end

  def direkt_enthaltene_vertraege
    unless @direkt_enthaltene_vertraege then
      self.direkt_enthaltene_vertraege = (Dir[self.pfad_dateimuster].select do |n|
         n !~ /(^|\/)[#.$][^\/]+/
      end.size)
      trc_temp :Excel_direkte=, @direkt_enthaltene_vertraege
    end
    @direkt_enthaltene_vertraege
  end

  def inhalt
    prefixe = excel_zugriff::MOEGLICHE_REFPREFIXE.join("|")
    liste = Dir[self.pfad_dateimuster].map { |ganzername|
      ganzername.sub(/^(.*\/)(#{prefixe})([^\/]+\.xls)$/, '\1\3')
    }.uniq.compact.sort

    liste.map { |ganzername|

      name = File.basename(ganzername)

      trc_temp :inhalt_zeile, name
      [IMGIDX_EXCEL, name] #, groesse, ref] #, ges, fehl, naja ]
    }
  end

  def zusatz_inhalt_fuer_item(dateiname)
    ganzername = ordner_pfad + dateiname
    kbytes =
      if File.exist?(ganzername)
        "%d kb" % (File.size(ganzername) / 1024)
      else
        "---"
      end
    ref_datei = ExcelZugriff.finde_refdatei(ganzername)
    ref = ""
    if ref_datei
      ref = "./" unless ref_datei =~ /\/Referenzdaten\/[^\/]+$/
      ref += ref_datei.match(/\/[^\/]+$/)[0][1,1]
    end
    (super(dateiname) || {}).update({:kbytes=>kbytes, :ref=>ref})
  end

  def on_vertr_dbl_click(item)
    require 'models/system_fkt/excel_zugriff'
    voller_dateiname = ordner_pfad + item.caption
    ExcelZugriff.application.Visible = true
    mappe = ExcelZugriff.application.Workbooks.Open( voller_dateiname )
    ExcelZugriff.als_vorderstes_fenster
  end

  def self.vliste_menu
    allg_vliste_menu + [
      { :text=>'Selektierte Mappen und zugehörige Referenzmappen öffnen',
        :name=>:menu_excel_oeffnen,
        :proc=>method(:excel_mit_ref_oeffnen_clicked)},
      { :text => "In &neuer Excel-Instanz öffnen",
        :proc => proc do |item|
          ort = aktueller_ort
          ez = excel_zugriff.neu_erzeugtes_excel(:sichtbar=>true)
          ez.app.Workbooks.open(ort.ordner_pfad + item.caption)
        end},
      { :text=>'&Referenzmappen aus selektierten Mappen erstellen',
        :name=>:menu_excel_refmappen_clicked,
        :proc=>method(:excel_refmappen_aus_mappen)},
      { :text=>'Mappen mit ihren Referenzmappen &vergleichen',
        :name=>:menu_excel_vergleich,
        :proc=>method(:excel_vergleichen_clicked)}
    ]
  end

  def place_spezifisches_menu
    trc_info :excelbase_spez_menu_abruf
    require 'models/system_fkt/excel_zugriff'
    OrtExcelBase.spez_menu_cache {
        trc_temp :excel_spez_menu_create
        form.make_main_menu(:excel_menu,
          [{
            :text=>'Excel',
            :name=>:menu_excel,
            :menu=>[{:text=>'Selektierte Mappen und zugehörige Referenz öffnen',
                     :name=>:menu_excel_oeffnen,
                     :proc=>self.class.method(:excel_mit_ref_oeffnen_clicked)},
                    {:text=>'&Referenzmappen aus selektierten Mappen erstellen',
                     :name=>:menu_excel_refmappen_clicked,
                     :proc=>self.class.method(:excel_refmappen_aus_mappen)},
                    {:text=>'Mappen mit ihrer Referenz &vergleichen',
                     :name=>:menu_excel_vergleich,
                     :proc=>method(:excel_vergleichen_clicked)}]
          }]
        )
    }
  end

  def self.excel_mit_ref_oeffnen_clicked(sender)
    require 'models/system_fkt/excel_zugriff'

    pfad = aktueller_ort.ordner_pfad
    ExcelZugriff.als_vorderstes_fenster
    @@controller.form.vliste.items.each {|item|
      if item.selected
        begin
          ExcelZugriff.oeffne_mappe_mit_refmappe(pfad + item.caption)
        rescue
          trc_aktuellen_error('excel_mit_ref_oeffnen_clicked fehlgeschlagen')
          @@controller.form.konsole.meldung "Konnte Mappe nicht öffnen: " + $!
        end
      end
    }
  end

  def self.excel_refmappen_aus_mappen(sender)

    pfad = aktueller_ort.ordner_pfad
    antwort = Phi::message_dlg(<<EOT,
Fuer jede ausgewaehlte Arbeitsmappe wird jetzt eine Kopie der Werte als Referenz erzeugt.
Die Referenz-Mappen werden im Unterordner "Referenzdaten" abgelegt,
dem Mappennamen wird ein "#{ExcelZugriff::STD_REFPREFIX}" vorangestellt.
Schon bestehende Mappen werden ueberschrieben.

Fortfahren?
EOT
               Phi::MT_CONFIRMATION,
               [Phi::MB_OK, Phi::MB_CANCEL], 0)
    return unless antwort == Phi::MR_OK
    exl = excel_zugriff.aktive_oder_neue_instanz
    exl.visible = true


    vetr_listbox = @@controller.form.vliste
    ungespeicherte_mappen = vetr_listbox.items.map do |item|
      next unless item.selected
      mappe = (exl.app.Workbooks(item.caption) rescue nil)
      mappe if mappe and not mappe.Saved
    end.compact

    if not ungespeicherte_mappen.empty? then
      antwort = Phi::message_dlg(<<EOT,
Beim Erstellen der Referenz-Kopien werden die Original-Mappen geschlossen.
Die Arbeitsmappe(n)
  #{ungespeicherte_mappen.map {|m| m.Name}.join("\n  ")}
enthalten ungespeicherte Änderungen.
Soll(en) die Mappe(n) vor dem Schließen gespeichert werden?
EOT
         Phi::MT_CONFIRMATION,
         [Phi::MB_YES, Phi::MB_NO, Phi::MB_CANCEL], 0
      )
      return if antwort == Phi::MR_CANCEL
      ungespeicherte_mappen.each do |mappe|
        case antwort
        when Phi::MR_YES then mappe.Save
        when Phi::MR_NO  then mappe.Saved = true
        end
      end
    end # unless

    ref_ordner = pfad + "Referenzdaten"
    Dir.mkdir(ref_ordner) unless File.exist?(ref_ordner)

    vetr_listbox.items.each do|item|
      if item.selected
        begin
          trc_temp :dateiname, pfad + item.caption
          mappe = exl.oeffne_oder_aktiviere( pfad + item.caption, :raise )
        rescue
          trc_aktuellen_error "oeffnen_bzw_aktivieren"
          @@controller.konsole.meldung "Mappe " + item.caption + " nicht gefunden."
          next
        end

        fn = mappe.FullName.gsub("\\","/")

        begin
          ref_fn = ExcelZugriff::refmappe_aus_mappenwerten(fn)
        rescue
          trc_aktuellen_error "refmappe_aus_mappenwerten"
          @@controller.konsole.meldung "#{$!} -- Ignoriert."
          next
        end

        @@controller.konsole.meldung "ReferenzMappe " + ref_fn + " erstellt."
        item.setze_inhalt(:ref=>ExcelZugriff::STD_REFPREFIX) #if listitem
        item.update

      end
    end
  end

  def self.excel_vergleichen_clicked(sender)
    aktueller_ort.excel_vergleichen_clicked(sender)
  end

  def excel_vergleichen_clicked(sender)
    if not self.equal?(aktueller_ort)
      aktueller_ort.excel_vergleichen_clicked(sender)
      return
    end
    mappen_namen = form.vliste.items.map { |item|
        item.caption if item.selected
      }.compact

    starte_dienst(Dienstherr_ExcelVergleich, mappen_namen)
  end

  def bewege_nach(ziel_pfad, option=nil)
    #return unless super
    require 'models/system_fkt/excel_zugriff'
    require 'ftools'
    zielname = ziel_pfad.gsub(/<Excel-Dateien>/,"").chomp("/")+"/"
    #return if File.exist?(neuer_name)

    trc_info :nach, zielname
    konsole.meldung ""


    arten = case option
      when :ref  then %w[/Referenzdaten/]
      when :work then %w[/]
      else            %w[/ /Referenzdaten/]
    end

    anzahl_kopierte_dateien = 0
    anzahl_kopierte_vertraege = 0
    arten.each do |x|
      (Dir[pfad_dateimuster.sub(/\/([^\/]*)$/, x+'\1')] + [persist_proxy.dateiname]).
          each do |dateiname|
        quelle = dateiname
        ziel   =  zielname.sub(/\/([^\/]*)$/, x+'\1') #+File.basename(quelle)
        trc_temp :qz, [quelle, ziel]
        begin
          if File.exist?(quelle)
            Phi::force_dirs((ziel))
            File.copy quelle, ziel #.chomp("/")
            anzahl_kopierte_dateien += 1
            anzahl_kopierte_vertraege += 1 if dateiname !~ /\/Referenzdaten\/|\.ats/
          end
        rescue
          trc_aktuellen_error "Kopieren fehlgeschlagen q,z= #{[quelle, ziel].inspect}", 3
        end
      end
    end
    super
    konsole.meldung "\n#{anzahl_kopierte_vertraege} Verträge (#{anzahl_kopierte_dateien} Dateien) kopiert."
  end
end

#**********************************************************

class OrtExlDat < OrtExcelBase
  def ordner_pfad
    erg = File.dirname(anzeige_pfad)+"/" # d.h. ohne "<Excel-Dateien>"
    erg
  end

  def physisch_loeschen
    return unless super
    require 'models/system_fkt/excel_zugriff'
    untergeordnete.each do |dateiname|
      begin
        File.delete(dateiname) if File.exist?(dateiname)
      rescue
        trc_info "Löschen fehlgeschlagen: ", dateiname
      end
      refdat = ExcelZugriff.finde_refdatei(dateiname)
      begin
        File.delete(refdat) if refdat
      rescue
        trc_info "Refdatei-Löschen fehlgeschlagen: ", refdat
      end
    end
    begin Dir.rm_rf(ordner_pfad+"Referenzdaten") rescue nil end
    self.komplett_aufloesen_sichtbar
  end

end


class OrtExlOrd < OrtExcelBase
  def ordner_pfad
    erg = super
#    trc_temp :OrtExlOrd_ornerpfad, erg
    erg
  end
  def physisch_loeschen
    return unless super
    Dir.rm_rf(self.ordner_pfad)
    self.komplett_aufloesen_sichtbar
  end

end

end # if not defined?(OrtExcelBase)

if __FILE__ == $0
  durchlaufe_unittests($0)
end
