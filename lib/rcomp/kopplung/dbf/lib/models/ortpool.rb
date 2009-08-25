
require 'schmiedebasis'

if not defined? OrtPool then

  def dateiname_zu_anzeigepfad(dateiname)
    (
      File.dirname(dateiname) +"/"+ case File.basename(dateiname)
      when /^(gz|rk|st|v[abdfkptv])[bs](.*)\.dbf$/i
        "<#{$2}>"
      when /^(.+\$\.xls)$/i
        "<#{$1}>"
      when /^.+\.xls$/i
        '<Excel-Dateien>'
      when /^gftest32\.exe$/i
        '<gftest32.exe>'
      else
        return nil
      end
    ).sub(/^\.\//, '')
  end

  class OrtPool
    def initialize(speicher_pfad)
      @speicher_pfad = speicher_pfad
      @poolhash = Hash.new
    end

    def entferne(pfad)
      if pfad.is_a? String then
        @poolhash.delete(normalisiere_pfad(pfad).downcase)
      end
    end
    
    def dateipfad_zu_ort(dateipfad, vater_vorgabe = :selbst_als_wurzel_wenn_keine_vorfahren)
      pfad_zu_ort(dateiname_zu_anzeigepfad(dateipfad), vater_vorgabe)
    end
    
    # :selbst_als_wurzel_wenn_keine_vorfahren
    # :vater_finden_sonst_nil
    def pfad_zu_ort(pfad, vater_vorgabe = :selbst_als_wurzel_wenn_keine_vorfahren)
      return nil if not pfad
      return nil if pfad == ""
      npfad = normalisiere_pfad(pfad)
      #trc_info :npfad_vater, [npfad, vater_vorgabe]
      hashpfad = npfad.downcase
      erg = if @poolhash.has_key?(hashpfad) and @poolhash[hashpfad] then
        @poolhash[hashpfad]
      else
        vater = if vater_vorgabe.is_a? Symbol then
          # also suchen wir den Vater den Pfad aufwärts
          vaterpfad = File.dirname(npfad)
          if vaterpfad.size >= npfad.size then # sicheres Zeichen für Ende des Pfads
            nil
          else
            pfad_zu_ort(vaterpfad, :vater_finden_sonst_nil)
          end
        else
          # check ob pfad enthalten
          vater_vorgabe
        end
        #trc_temp :vater,  vater

        if vater or vater_vorgabe == :selbst_als_wurzel_wenn_keine_vorfahren then
          @poolhash[hashpfad] = Ort_Haupt.neu($ats, vater, npfad)
        else
          nil
        end
      end
      #trc_info :erg, erg
      erg
    end


  end

end
