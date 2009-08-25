# encoding: iso-8859-15

unless defined? DateiAufraeumer then
  
  DATEI_AUFRAEUMER_DEFAULT_HASH = {
    :ordner => ".",
    :einzel_muster => /.*\.log/,
    :paket_muster => /paket_.*\.log/,
    :einzel_zu_paket_name => proc {|en| "paket_" + en},
    :einzel_fixiert_anzahl => 5, 
    :einzel_freigabe_ab_anzahl => 30, 
    :einzel_freigabe_ab_bytes => 2_000_000,      
    :paket_freigabe_ab_bytes => 2_000_000,      
    :paket_datei_bytes => 400_000     
  }
  
  class DateiAufraeumer < Struct.new(* DATEI_AUFRAEUMER_DEFAULT_HASH.keys)
    
    def initialize(konfig)
      super
      setze_konfig(DATEI_AUFRAEUMER_DEFAULT_HASH)
      setze_konfig(konfig)
      @aktions_queue = [
        "einzel_dateien_sammeln",
        "einzel_dateien_alte_loeschen",
        "paket_dateien_sammeln_und_alte_loeschen",
        "einzel_dateien_zusammenfassen"
      ]
    end 
      
    def setze_konfig(konfig_hash)
      konfig_hash.each {|k,v| self[k] = konfig_hash[k]} 
    end  
    
    def next_aufraeum_aktion
      if @aktions_queue.empty? then
        trc_info :keine_aktionen_mehr
        "fertig"
      else
        next_method = @aktions_queue.shift
        trc_info :next_aktion, next_method
        bericht = method(next_method).call
        "#{next_method}: #{bericht}"
      end
    end
    
    def aktions_beginn(symbol)
      trc_temp symbol
    end  
    
    def aktions_ende(symbol)
      trc_temp symbol
      @aktions_queue.delete(symbol)
    end  
    
    def einzel_dateien_sammeln
      aktions_beginn :einzel_dateien_sammeln
      sortierte_einzel_dateien = Dir[self.ordner + "/*"].select do |voller_dateiname|
        self.einzel_muster =~ File.basename(voller_dateiname) 
      end.sort_by do |voller_dateiname|
        File.mtime voller_dateiname
      end.reverse
      trc_info :anz_einzel_dateien, sortierte_einzel_dateien.size
      trc_info "5_einzel_dateien", sortierte_einzel_dateien.last(5)
      
      aktions_ende :einzel_dateien_sammeln
      @sortierte_einzel_dateien = sortierte_einzel_dateien
      ges_bytes = sortierte_einzel_dateien.inject(0) {|s,d| s + File.size(d)}
      
      "#{sortierte_einzel_dateien.size} dateien mit #{ges_bytes} bytes"      
    end
  
    def einzel_dateien_alte_loeschen
      aktions_beginn :einzel_dateien_alte_loeschen
      erst_veraendern_ab = @sortierte_einzel_dateien.each_with_index do |voller_dateiname, idx|
        next if idx < self.einzel_fixiert_anzahl
        trc_temp "voller_dateiname", (voller_dateiname)
        trc_temp "File.mtime(voller_dateiname)", File.mtime(voller_dateiname)
        mtime = File.mtime(voller_dateiname)
        File.basename(voller_dateiname) =~ /_(\d\d)-(\d\d)_(\d\d)(\d\d)_/ # inf konfig rausziehen    
        mtime = Time.local(Time.now.year, $1, $2, $3, $4)
        trc_temp "mtime", mtime
        next if mtime > Time.now - 2 * 24 * 60 * 60
        break idx
      end
      if erst_veraendern_ab == @sortierte_einzel_dateien then
        # break nicht aufgerufen --> nichts verändern
        @sortierte_einzel_dateien = []
      else
        @sortierte_einzel_dateien = @sortierte_einzel_dateien[erst_veraendern_ab .. -1]
      end
  
      bericht = loesche_alte_dateien!(@sortierte_einzel_dateien, 
                            self.einzel_freigabe_ab_anzahl, 
                            self.einzel_freigabe_ab_bytes)
      aktions_ende :einzel_dateien_alte_loeschen
      bericht
    end
    
    def paket_dateien_sammeln_und_alte_loeschen
      aktions_beginn :paket_dateien_sammeln_und_alte_loeschen
      paket_dateien_hash = {}
      Dir[self.ordner + "/*"].each do |voller_dateiname|
        match = self.paket_muster.match(File.basename(voller_dateiname))
        if match then
          alles, jahr, mon, tag, std, min = match.to_a.map {|str| str.to_i}
          z = Time.mktime(jahr, mon, tag, std, min)
          paket_dateien_hash[z] = voller_dateiname
        end
      end
  
      sortierte_paket_dateien = paket_dateien_hash.sort_by do |zeit, voller_dateiname|
        zeit
      end.map do |zeit, voller_dateiname|
        voller_dateiname
      end.reverse
  
      bericht = loesche_alte_dateien!(sortierte_paket_dateien, 
                            2, 
                            self.paket_freigabe_ab_bytes)
                            
      aktions_ende :paket_dateien_sammeln_und_alte_loeschen
      @sortierte_paket_dateien = sortierte_paket_dateien
      bericht 
    end
    
    
    def einzel_dateien_zusammenfassen
      aktions_beginn :einzel_dateien_zusammenfassen
      limit = self.paket_datei_bytes
      bisher_gesammelter_inhalt = nil  
      akt_einzel = nil 
      offenes_paket = nil
      gesammelte_einzel_namen = nil
      if @sortierte_einzel_dateien and @sortierte_paket_dateien then
        dliste = [@sortierte_einzel_dateien, @sortierte_paket_dateien]
        vorher = dliste.map{|d| d.size}
        catch :fertig do
          loop do
            begin
              loop do
                throw :fertig if @sortierte_einzel_dateien.empty?
                akt_einzel = @sortierte_einzel_dateien.pop
                if not bisher_gesammelter_inhalt then # hier wird initialisiert
                  gesammelte_einzel_namen = []
                  offenes_paket = @sortierte_paket_dateien.first
                  if offenes_paket and File.size(offenes_paket) + File.size(akt_einzel) > limit then
                    offenes_paket = nil # schon zu groß, gleich wieder schließen
                  end
                  bisher_gesammelter_inhalt = (offenes_paket ? File.read(offenes_paket) : "")
                end
          
                bisher_gesammelter_inhalt += File.read(akt_einzel)
                gesammelte_einzel_namen << akt_einzel
          
                break if bisher_gesammelter_inhalt.size + File.size(akt_einzel) > limit
              end
            ensure
              if bisher_gesammelter_inhalt and bisher_gesammelter_inhalt > "" then
              # jetzt schreiben
                paket_name = self.einzel_zu_paket_name.call(akt_einzel)
                File.open(paket_name, "w") do |schreibdatei|
                  schreibdatei.puts bisher_gesammelter_inhalt
                end
                File.delete(offenes_paket) if offenes_paket # wir haben ja jetzt eine neue
                File.delete(* gesammelte_einzel_namen)
                @sortierte_paket_dateien.unshift paket_name
                bisher_gesammelter_inhalt = nil
              end          
            end
          end
        end # catch
        nachher = dliste.map{|d| d.size}
        bericht = "einzel/paket vorher:#{vorher.join('/')}, nachher:#{nachher.join('/')}"
      else
        bericht = "Fehler: @sortierte_einzel_dateien und @sortierte_paket_dateien müssen belegt sein"
      end
      aktions_ende :einzel_dateien_zusammenfassen
      bericht
    end
  
  #private
      
    def loesche_alte_dateien!(sortierte_datei_liste, behalte_anzahl, behalte_groesze)
      #return ##+# deaktiviert
      gesamte_groesze_bisher = 0
      grenz_idx = sortierte_datei_liste.each_with_index do |voller_dateiname, idx|
        break idx if (idx+1 > behalte_anzahl and 
                     gesamte_groesze_bisher >= behalte_groesze)
        gesamte_groesze_bisher += File.size(voller_dateiname)
      end
  
      if grenz_idx != sortierte_datei_liste then # d.h. break aufgerufen, ansonsten nichts zu löschen übrig
        File.delete(*sortierte_datei_liste.slice!(grenz_idx .. -1))
      else
        grenz_idx = "alle #{sortierte_datei_liste.size}"
      end  
      bericht = "behalte #{grenz_idx} Dateien mit #{gesamte_groesze_bisher} bytes"
    end
  
  end # class DateiAufraeumer
end

if __FILE__ == $0
  durchlaufe_unittests($0)
end

