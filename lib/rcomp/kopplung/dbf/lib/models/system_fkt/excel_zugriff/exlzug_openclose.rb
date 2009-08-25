
require 'models/system_fkt/excel_zugriff/exlzug_basis'

if not defined? ExcelErrorOpen then

class ExcelZugriff

  MOEGLICHE_REFORDNER  = ["", "Referenzdaten/"]
  MOEGLICHE_REFPREFIXE = ["$", "#", "§", "Kopie von "]
  STD_REFPREFIX = MOEGLICHE_REFPREFIXE[0]

  def self.finde_refdatei(voller_dateiname)
    pfad     = File.dirname(voller_dateiname)
    basename = File.basename(voller_dateiname)
    MOEGLICHE_REFPREFIXE.each { |prefix|
      MOEGLICHE_REFORDNER.each { |ordner|
        ref_dateiname = pfad + '/' + ordner  + prefix + basename
        return ref_dateiname if File.exist?(ref_dateiname)
      }
    }
    nil
  end

  # Für Dateinamen, die sich nicht ï¿½ffnen lassen, wird der Block aufgerufen.
  # öffnet eine Arbeitsmappe und deren Referenzdatei
  def self.oeffne_mappe_mit_refmappe(voller_mappenname)

    if not File.exist?(voller_mappenname.tr("\\","/"))
      raise ExcelErrorOpenDateiNichtGefunden, "Datei #{voller_mappenname} nicht gefunden"
    end

    if KONFIG.opts[:ExcelImmerWiederNeuStarten]
    #  self.beenden
    end

    sichtbarkeit_reparieren if self.sichtbar

    ref_dateiname = finde_refdatei(voller_mappenname)

    aktion_falls_ungespeicherte_mappe = :raise
    
    refmappe = if ref_dateiname then
      oeffne_oder_aktiviere(ref_dateiname, aktion_falls_ungespeicherte_mappe)
    end
    mappe = oeffne_oder_aktiviere(voller_mappenname, aktion_falls_ungespeicherte_mappe)
    
    raise "Es wurde keine Referenz-Mappe gefunden" if not ref_dateiname 

    [mappe, refmappe]
  end

  # aktion_falls_schon_existiert kann sein: :raise, :overwrite oder :excel
  def self.speichere_kontrolliert(mappe, dateiname, aktion_falls_schon_existiert)
    datname = dateiname.gsub("//","/")
    begin
      mappe_ist_noch_zu_speichern = true
      if File.exist?(datname) then
        case aktion_falls_schon_existiert
        when :overwrite  then
          if mappe.Name == File.basename(datname) then
            mappen_fullname = mappe.FullName.gsub(/[\/\\]/, "/")
            if mappen_fullname == datname then
              mappe.Save
              mappe_ist_noch_zu_speichern = false
            else
              # alles ok (kann nicht im Weg stehen, kann gelöscht werden)
              trc_info :verschiedene_dateien, [mappe.FullName, datname]
            end
          else
            mappe_im_weg = begin
              mappe.Application.Workbooks(File.basename(datname)) 
            rescue WIN32OLERuntimeError
              nil
            end
            if mappe_im_weg then 
              if mappe_im_weg.Saved then 
                mappe_im_weg.Close   
              else
                raise "Mappe kann nicht gespeichert werden, da eine andere mit gleichem Namen geöffnet ist. (ziel: #{datname}"
              end
            end
            
          end
          File.delete(datname) if mappe_ist_noch_zu_speichern
        when :excel  then #nix
        when :raise
          raise ExcelErrorSaveMappeStehtImWeg, "Mappe existiert bereits: #{datname}"
        else
          raise ExcelErrorSave, "Bug: Ungültige Option (#{aktion_falls_schon_existiert.inspect})"
        end
      end
      mappe.SaveAs(datname) if mappe_ist_noch_zu_speichern
    rescue WIN32OLERuntimeError
      trc_aktuellen_error "mappe.SaveAs(#{datname})"
      raise "Konnte die Arbeitsmappe nicht als #{datname} speichern"
    end
  end

  # aktion_falls_ungespeicherte_mappe kann sein :raise, :forget, :accept oder :readonly
  def self.oeffne_oder_aktiviere(voller_dateiname, aktion_falls_ungespeicherte_mappe)
    aktive_oder_neue_instanz.oeffne_oder_aktiviere(voller_dateiname, aktion_falls_ungespeicherte_mappe)
  end

  # Wenn die Datei nicht existiert, wird eine Exception ausgelöst
  # Wenn nil als Deiteiname angegeben wird, passiert nichts
  # aktion_falls_ungespeicherte_mappe kann sein :raise, :forget, :accept oder :readonly
  def oeffne_oder_aktiviere(voller_dateiname, aktion_falls_ungespeicherte_mappe)
    return nil unless voller_dateiname
    if not File.exist?(voller_dateiname)
      raise ExcelErrorOpenDateiNichtGefunden, "Datei #{voller_dateiname} nicht gefunden"
    end
    trc_info :oeffne_oder_aktiviere, [aktion_falls_ungespeicherte_mappe, voller_dateiname]
    mappe_schon_geoeffnet = begin
      mappe = @app.Workbooks(File.basename(voller_dateiname))
      trc_temp :oeff_o_akt_existierend, mappe.FullName
      true
    rescue WIN32OLERuntimeError
      false
    end

    if mappe_schon_geoeffnet then
      # wirklich die richtige Mappe?
      ist_richtige_mappe = mappe.FullName.tr("\\","/").downcase == voller_dateiname.tr("\\","/").downcase
      # eine Mappe, die den gleichen Namen hat, aber in einem andern Ordner steht
      # Sie ist im Weg, denn Excel kann die neue Mappe nicht gleichzeitig öffnen

      if aktion_falls_ungespeicherte_mappe == :readonly then
        if ist_richtige_mappe then
          trc_temp :def_freigabe, :nix
          bei_freigabe_proc = proc do 
            1 # nichts tun, die Mappe bleibt also offen
          end
        else
          neues_excel = self.class.neu_erzeugtes_excel
          mappe = neues_excel.oeffne_oder_aktiviere(voller_dateiname, :raise)
          ist_richtige_mappe = true
          trc_temp :def_freigabe, :neues_excel
          bei_freigabe_proc = proc do
            mappe.Close #rescue nil
            neues_excel.beenden
          end
        end                  
        ec = class << mappe; self; end
        ec.send(:define_method, :bei_freigabe) do
          trc_temp :freiganf
          trc_temp :freigeben, (mappe.Name rescue nil)          
          bei_freigabe_proc.call
        end
          
        
      elsif not mappe.Saved then
        case aktion_falls_ungespeicherte_mappe
        when :raise then
          if ist_richtige_mappe
            raise ExcelErrorOpenMappeSelbstUngespeichert, "Die Mappe #{mappe.Name} besitzt ungespeicherte Änderungen"
          else
            raise ExcelErrorOpenMappeStehtImWeg, "Es ist bereits eine Mappe mit demselben Namen geöffnet, und diese besitzt ungespeicherte ï¿½nderungen (#{mappe.FullName})"
          end
          
        when :accept then # wenn es die falsche Mappe ist, müssem wir trotzdem meckern
          if not ist_richtige_mappe
            raise ExcelErrorOpenMappeStehtImWeg, "Es ist bereits eine Mappe mit demselben Namen geöffnet, und diese besitzt ungespeicherte ï¿½nderungen (#{mappe.FullName})"
          end
          
        when :forget then 
          mappe_schon_geoeffnet = false
          mappe.Close(false)         
        else
          raise ExcelErrorFalscherParameter, "Ungültige Aktionsart (#{aktion_falls_ungespeicherte_mappe})"
        end
      end

    end

    if not mappe_schon_geoeffnet then
      begin
        mappe = @app.Workbooks.Open(voller_dateiname)
        if aktion_falls_ungespeicherte_mappe == :readonly then
          def mappe.bei_freigabe
            trc_temp :freigeben, (self.Name rescue nil)
            self.Close
          end
          trc_temp :def_freigabe, :close
        end
      rescue WIN32OLERuntimeError
        raise ExcelErrorOpenDateiNichtGefunden, $!.to_s
      end
#      trace :oeff_o_akt_neu, mappe.FullName
    end

    if mappe.nil?
      raise ExcelErrorOpenExcelNichtsGeliefert, "Excel hat die Arbeitsmappe nicht geöffnet, aber auch keinen Fehler gemeldet (Dateiname=#{voller_dateiname})"
    end
    trc_temp :oeff_o_akt_ende, [mappe.respond_to?(:bei_freigabe), mappe.FullName]
    mappe.Activate
    mappe
  end
  
  def self.lese_aus_mappe(voller_mappename)
    mappe = aktive_oder_neue_instanz.oeffne_oder_aktiviere(voller_mappename, :readonly)
    begin      
      erg = yield mappe
    ensure
      mappe.bei_freigabe if mappe.respond_to? :bei_freigabe
    end     
    erg 
  end

  def self.multi_oeffnen(ort, eintraege)
    trc_info :eintraege, eintraege
    fehlermappen = []
    eintraege.each { |(datei, hitem)|
      mappe, refmappe = oeffne_eine(ort, datei) { |gescheiterter_dateiname|
        trc_info "nicht geklappt: ", gescheiterter_dateiname
        fehlermappen << gescheiterter_dateiname
      }
      if mappe and block_given?
        yield  mappe, refmappe, hitem

        if KONFIG.opts[:MappeSpeichernBeimVgl]
          mappe.Save
        else
          mappe.Saved = true
        end
        mappe.Close
        refmappe.Saved = true
        refmappe.Close
        mappe = nil
        refmappe = nil
      end
    }
    schlieszen
    #beenden
    fehlermappen
  end

  def self.oeffne_vorg_bibbliothek_falls_gefordert
    if KONFIG.opts[:VorgBibliothekNutzen]
      begin
        bibname = KONFIG.opts[:VorgegebeneBibliothek]
        trc_info :konfig_vorg_bib, bibname
        if bibname > ""
          oeffne_oder_aktiviere(bibname, :raise)
        end
      rescue WIN32OLERuntimeError
        trc_aktuellen_error :vorgabe_nicht_geoeffnet
        schlieszen(:raise)
        retry
        #raise "Vorgegebene Exstar-Bibliothek nicht geöffnet"
      end
    end
  end

  def self.existiert_vorgegebene_bibbliothek?(bib_dateiname = nil)
    return true unless KONFIG.opts[:VorgBibliothekNutzen]

    bib_dateiname ||= KONFIG.opts[:VorgegebeneBibliothek] 
    File.exist?(bib_dateiname)
  end

  def self.refmappe_aus_mappenwerten(fn, aktion_falls_ungespeicherte_mappe=:raise)
    mappe = oeffne_oder_aktiviere(fn, aktion_falls_ungespeicherte_mappe)
    ref_fn = File.dirname(fn) + "/Referenzdaten/" + ExcelZugriff::STD_REFPREFIX + File.basename(fn)
    begin
      xl_app.Workbooks(File.basename(ref_fn)).Close(false)
    rescue
    end

    begin
      speichere_kontrolliert(mappe, ref_fn, :overwrite)
    rescue
      trc_aktuellen_error "refmappe speichern"
      raise "Problem beim Speichern von Mappe " + ref_fn
    end
    
    nurnoch_werte_behalten(mappe)
    
    mappe.Save
    mappe.Close
    ref_fn
  end

  def self.nurnoch_werte_behalten(mappe)
    alle_blattnummern = (1 .. mappe.Sheets.Count).to_a
    trc_info :alle_blattnummern, alle_blattnummern
    xl_app = mappe.Application
    mappe.Sheets(alle_blattnummern).Select
    xl_app.Cells.Select
    xl_app.Selection.Copy
    xl_app.Selection.PasteSpecial "Paste"=>XLPasteValues
    mappe.Sheets(1).Select # Gruppenmarkierung aufheben    
  end


  
end


class ExcelErrorOpen < ExcelError
end

class ExcelErrorOpenDateiNichtGefunden < ExcelErrorOpen
end

class ExcelErrorOpenExcelNichtsGeliefert < ExcelErrorOpen
end

class ExcelErrorOpenMappeUngespeichert < ExcelErrorOpen
end

class ExcelErrorOpenMappeStehtImWeg < ExcelErrorOpenMappeUngespeichert
end

class ExcelErrorOpenMappeSelbstUngespeichert < ExcelErrorOpenMappeUngespeichert
end


class ExcelErrorSave < ExcelError
end
class ExcelErrorSaveMappeStehtImWeg < ExcelErrorSave
end

end # if not defined? ...



if __FILE__ == $0 then
  durchlaufe_unittests($0)
end
