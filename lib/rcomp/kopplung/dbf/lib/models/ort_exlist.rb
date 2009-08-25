

if not defined?(OrtExList)

require 'models/ortpool'
require 'models/ort_haupt'



#*************************************************************


class OrtExList < Ort_Haupt

  def basename
    name. sub(/^</,""). sub(/>$/,"")
  end
  
  def kanonischer_pfad
    anzeige_pfad.gsub(/[><]/, "").chomp("/")
  end

  def system_pfad
    kanonischer_pfad
  end
  
  def ordner_pfad
    File.dirname(anzeige_pfad)+"/"  # also ohne den Base-Namen
  end
  
  def generiere_unterorte(auswahl_pfad = nil)
    gesamtanz_vertraege
  end

  def spalten_defs
    [
      {:text => "VsNr", :width => 110}
      #{:name=>:kbytes,    :text => "Groesze", :width=>60, :align => Phi::TA_RIGHT_JUSTIFY},
      #{:name=>:ref,       :text => "Ref", :width => 33, :align => Phi::TA_CENTER}
    ] +
      super
  end

  def self.image_index
    IMGIDX_EXCEL
  end

  def pfad_dateimuster
    kanonischer_pfad
  end

#  def base_datei_regex
 #   @@base_datei_regex ||= /\/[^#{excel_zugriff::MOEGLICHE_REFPREFIXE[0,3].join("")}][^\/]*\.xls$/i
  #end

  def untergeordnete
    trc_temp :exlist_untg_anzeige_pfad, anzeige_pfad
    
    #require 'models/system_fkt/excel_direkt'
    require 'parseexcel'
    
    if not @untergeordnete then
      
      vsnrn = []
      
      mappe = Spreadsheet::ParseExcel.parse(kanonischer_pfad)
      blatt = mappe.worksheet(1)
      spalten_nr = blatt.row(0).index("VSNR")
      if not spalten_nr then
        spalten_nr = blatt.row(1).each_with_index {|z, i| break i if z}
        spalten_nr = 0 unless spalten_nr.is_a? Numeric
      end  
      #Range("VSNR")
      blatt.each(1) do |zeile|
        break unless zeile
        begin
          zelle = zeile[spalten_nr]
          if zelle then
		        trc_temp :typ, zelle.type
            vsnr = zelle.to_s("ISO-8859-1")
            vsnr.sub!(/\.0$/, "") # #*#HACK!!
		        vsnrn << vsnr if vsnr !~ /XXXX/
          end
        rescue
          trc_aktuellen_error "parsexl (zeil:#{zelle.value})"
        end
      end
      @untergeordnete = vsnrn.uniq.sort
    end
    
    @untergeordnete
  end

  def direkt_enthaltene_vertraege
    unless @direkt_enthaltene_vertraege then
      self.direkt_enthaltene_vertraege = untergeordnete.size
    end
    @direkt_enthaltene_vertraege
  end

  def inhalt
    untergeordnete.map { |vsnr|
      trc_temp :inhalt_zeile, vsnr
      [IMGIDX_EXCEL, vsnr] #, groesse, ref] #, ges, fehl, naja ]
    }
  end

  def zusatz_inhalt_fuer_item(dateiname)
    return super 
    
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
    voller_dateiname = kanonischer_pfad 
    ExcelZugriff.application.Visible = true
    mappe = ExcelZugriff.oeffne_oder_aktiviere( voller_dateiname, :accept )
    mappe.Activate
    ExcelZugriff.als_vorderstes_fenster
    trc_info :item_caption, item.caption
    if mappe then
      blatt = mappe.Activesheet #.Cells
      ziel = blatt.Cells.Find("What"  => item.caption, 
                              "After" => blatt.Cells(1,1), 
                              "LookIn"=> XLValues,  
                            #  "LookAt"=> XLPart, 
                            #  "SearchOrder", 
                              "MatchByte" => false)
      if ziel
        begin
          alte_zelle = mappe.Application.ActiveCell rescue nil
          zeile = ziel.EntireRow
          zeile.Select
          blatt.Cells(ziel.Row, alte_zelle.Column). Activate if alte_zelle
          trc_info :aktiviert_zeile, ziel.Row 
        rescue
          trc_aktuellen_error "zeile mit #{item.caption} nicht aktivierbar", 7
        end
      else
        trc_info "zeile nicht gefunden"
      end
    else 
      trc_hinweis :mappe_ist_nil, voller_dateiname
    end
  end

  def self.vliste_menu
    allg_vliste_menu + [
      { :text=>'Mappe öffnen und zu aktueller Position springen',
        :proc=>method(:excel_mit_ref_oeffnen_clicked)},
      { :text => "In &neuer Excel-Instanz öffnen",
        :proc => proc do |item|
          ort = aktueller_ort
          ez = excel_zugriff.neu_erzeugtes_excel(:sichtbar=>true)
          ez.app.Workbooks.open(ort.system_pfad )
          #+ item.caption
        end},
      { :text=>'Mappe mit in ihr gespeicherter Referenz-Information &vergleichen',
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
            :menu=>[{:text=>'Mappe öffnen und zu aktueller Position springen',
                     :name=>:menu_excel_oeffnen,
                     :proc=>self.class.method(:excel_mit_ref_oeffnen_clicked)},
                    {:text=>'Mappe mit in ihr gespeicherter Referenz-Information &vergleichen',
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
          on_vertr_dbl_click(item)
          #ExcelZugriff.oeffne_mappe_mit_refmappe(pfad+item.caption, :raise)
        rescue
          trc_aktuellen_error('excel_mit_ref_oeffnen_clicked fehlgeschlagen')
###          form.konsole.meldung "Konnte Mappe nicht ffnen: " + $!
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


end # if not defined?(OrtExcelBase)

if __FILE__ == $0
  durchlaufe_unittests($0)
end
