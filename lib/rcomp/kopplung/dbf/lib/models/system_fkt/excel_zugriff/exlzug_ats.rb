
require 'models/system_fkt/excel_zugriff/exlzug_openclose'

if not defined? ExstarError then


class ExcelZugriff

  def self.excel_vorlage_oeffnen(wandel_art)
    trc_temp :begin_oeffn_wa=, wandel_art
    wandel_art_str = (wandel_art.is_a?(WandelArt) ? wandel_art.bez : wandel_art.to_s) 
    begin
      xl = ExcelZugriff.application
      #   xl.Interactive = false
      #  ExcelZugriff.als_vorderstes_fenster

      trc_info :ORIG_DIRNAME, ORIG_DIRNAME
      ExcelZugriff.oeffne_vorg_bibbliothek_falls_gefordert

      trc_temp(:KONFIG_opts, KONFIG.opts) 

      vorlage_dateiname = KONFIG.opts["ExstarVorl_#{wandel_art_str}".to_sym]
      vorlage_dateiname = ORIG_DIRNAME + "/" + vorlage_dateiname if vorlage_dateiname !~ /\//
      trc_hinweis :vorl_datnam, vorlage_dateiname
      mappe = ExcelZugriff.oeffne_oder_aktiviere(vorlage_dateiname, :raise)
      # schon in obiger Funktion erledigt:
      # mappe.Activate
      begin
        xl.Calculation = -4135 
      rescue WIN32OLERuntimeError
      end
      ## xlCalculationManual

      trc_info :vorl_mappe, mappe

      vorlage_art = sicheres_lesen("ExstarVorlage_Art", "Die Art der Exstar-Vorlage konnte nicht ermittelt werden.")
      trc_hinweis :vorlage_art, vorlage_art

      if vorlage_art != wandel_art_str then
        raise "Die Excel-Vorlage passt nicht zu der gewählten Art der dbf-Umwandlung."
      end

      ats_mindest_version = sicheres_lesen("Testschmiede_MindestVersion")
      trc_hinweis :ats_mindest_version, ats_mindest_version

      if (Version.new(ats_mindest_version) <=> Ats::VERSION::OBJ) == 1 then
        raise "Die Exstar-Vorlage ist zu neu." +
          "\nSie erfordert Testschmiede-Version:  #{ats_mindest_version})" +
          "\nFolgende Vorlagedatei wurde benutzt: #{vorlage_dateiname}"
      end


      begin
        vorlage_module = Object.const_get( ("ExcelZuordnung"+vorlage_art).to_sym )
      rescue
        trc_aktuellen_error :vorlage_module
        raise "Keine Zuordnung zum Umwandeln der Exstar-Vorlage (Art: '#{vorlage_art}') gefunden"
      end
      #eingaben_blatt = mappe.Sheets("Eingaben")

      vorlage_version = sicheres_lesen("ExstarVorlage_Version", "Die Version der Exstar-Vorlage konnte nicht ermittelt werden")
      trc_hinweis :vorlage_version, vorlage_version

      if vorlage_version < vorlage_module::VORLAGE_MINDEST_VERSION then
        raise "Die Exstar-Vorlage ist zu alt. (Erforderliche Version: #{vorlage_module::VORLAGE_MINDEST_VERSION})" +
        "\nBenutzete Vorlage: Version:#{vorlage_version} Datei: #{vorlage_dateiname}"
      end

    rescue WIN32OLERuntimeError
      trc_aktuellen_error "Vorlage-Mappe nicht zu öffnen. (für #{vorlage_dateiname})"
      raise "Konnte die Vorlage-Arbeitsmappe nicht öffnen. (Name: #{vorlage_dateiname})"
    end
    mappe
  end

  

  def self.vergleiche_mappe(voller_mappenname)
    trc_info :voller_mappenname, voller_mappenname
    ist_neu_form = (voller_mappenname =~ /^(.+)\/(<?)([^\/]+\$\.xls)(>|\/|$)/i)
    voller_dateiname = ist_neu_form ? $1 + "/" + $3 : voller_mappenname 

    ref_dateiname = finde_refdatei(voller_dateiname)
    
    raise "Referenzdatei fehlt" unless ref_dateiname or ist_neu_form
            
    self.sichtbar = KONFIG.opts[:MappenAnzeigenBeiBatchjobs]
    trc_info :vergleiche_mappe_sichtbar, sichtbar
    if KONFIG.opts[:ExcelImmerWiederNeuStarten]
      self.beenden
      #sleep 2 #*# wieder entfernen
    end

    oeffne_vorg_bibbliothek_falls_gefordert

    mappe = oeffne_oder_aktiviere(voller_dateiname, :raise)
    begin
      if not ref_dateiname then
        blattnamen = []
        mappe.Worksheets.each { |blatt| blattnamen << blatt.Name }
        if blattnamen.find { |bn| bn =~ /-ref$/ } then
          refmappe = nil
        else  
          return "Keine Referenzinformation"
        end
      else
        refmappe = oeffne_oder_aktiviere(ref_dateiname, :raise)        
      end
      
      trc_info :mappe_refmappe, [mappe, refmappe]
      #ExcelZugriff.application.ScreenUpdating = false

      bibname, bibversion = ermittle_exstarkern_bibliothek
      

      vgl = lebende_verbundene_instanz.vergleichen_intern(mappe, refmappe, bibversion)

      # #*# experimentell:
      sichtbarkeit_reparieren if self.sichtbar

    ensure
      trc_info "Closing"
      if KONFIG.opts[:MappeSpeichernBeimVgl]
        mappe.Save
      else
        mappe.Saved = true
      end
      mappe.Close
      if refmappe then
        refmappe.Close(false)
      end
    end
    vgl
  end

  def self.ermittle_exstarkern_bibliothek
    bibs = []
    application.Workbooks.each { |m|
      bibs << m if m.Name =~ /^ExstarKern/
    }
    trc_info :bibs, bibs
    bib = bibs.last
    if bib
      bib.Activate
      bibname = bib.Name
      begin
        bibliotheksversion = application.Run("ExstarKern_Version")
      rescue WIN32OLERuntimeError
        bibliotheksversion = bibname
      end
    else
      trc_info :ExstarBibbliothek_nicht_da
      bibname = "keine_bibliothek"
      bibliotheksversion  = "---"
    end
    trc_info :bibliotheksversion , bibliotheksversion
    [bibname, bibliotheksversion]
  end

  def fehler_faenger(abfangen)
    #,                          :fehler_nicht_abfangen
    #, option=:fehler_abfangen)
    case option
      when :fehler_nicht_abfangen then fehler_abfangen = false
      when :fehler_abfangen       then fehler_abfangen = true
      else raise "Ungültige Option für ExcelZugriff.vergleichen: ''#{option}''"
    end

    if abfangen then
      begin
        yield
      rescue
      end
    else
      yield
    end
  end



  def vergleichen_intern(mappe, refmappe, bibname)
    trc_info :beginne_vergleich, [mappe.Name, refmappe && refmappe.Name]
    excel = mappe.Application
  
    begin
      makro_mappen_name = "E2Vergleich.xls"
      makro_mappe = excel.Workbooks(makro_mappen_name)
    rescue WIN32OLERuntimeError
      makro_mappe = excel.Workbooks.Open(ORIG_DIRNAME + "/" + makro_mappen_name)
    end
    makro_mappe.IsAddin = true
  
    t1 = Time.now
    gesamt_erg = VergleichsErgebnis.new
    #yield_to_events
  
    if KONFIG.opts[:BerechnenVorVgl]
      komplett_berechnen(mappe)
    end
    trc_temp :ws, mappe.Worksheets.class
  
    max_inte, interessantestes_blatt = -1, nil
    mappe.Worksheets.each do |blatt|
      #yield_to_events
      blatt.Activate
      begin
        if refmappe.nil? then
          refblatt = mappe.Worksheets(blatt.Name + "-ref")
        else
          refblatt = refmappe.Worksheets(blatt.Name)
        end
      rescue
        next
      end
      makro_mappe.activate
      #trace "vor-blattvgl"
      vergleich_opts = "NurFormeln:#{!!KONFIG.opts[:NurZellenMitFormelnVergleichen]}"
                                        
#      vgl_coll = 
      nur_namen = %w[RmAbsk LaufAZ VwkPr RatzusVwk StornoAb]
      vergleich_opts << "#DefPer:ref"
      #+# interessante Namen auswählen!
      #vergleich_opts << "#NurNamen:" + nur_namen.join("ï¿½")
      
      erg = excel.Run("BlattVergleichen", blatt, refblatt, vergleich_opts) # erg ist Array der Ergebnis-Anzahlen (indiziert nach ERGEBNIS_ARTEN)
      trc_info "nach-blattvgl", erg
  
      vgl_erg = VergleichsErgebnis.new
      vgl_erg.bib = bibname
      VergleichsErgebnis::ERGEBNIS_ARTEN.each_with_index {|abw_art, idx|
        vgl_erg[abw_art] = erg[idx]
      }
      inte = 10000*vgl_erg[:eaVollDaneben]     + 1000*vgl_erg[:eaXception] +
               100*vgl_erg[:eaUngenau,:eaNochOK] + 10*vgl_erg[:eaFastExakt] +
                   vgl_erg.summe
      if inte > max_inte
        max_inte               = inte
        interessantestes_blatt = blatt
      end
      gesamt_erg += vgl_erg
    end
    interessantestes_blatt.Activate if interessantestes_blatt
  
    t2 = Time.now
    trc_info "Vergeichsdauer:", t2-t1
    gesamt_erg
  end

  def komplett_berechnen(mappe)
    excel = mappe.Application
    mappe.Activate
    begin
      excel.Run("RUnit.xls" + "!AutoKomplettBerechnen")
    rescue WIN32OLERuntimeError
      trc_info "AutoKomplettBerechnen failed, mappe=", mappe.Name
      begin # Workaround gegen abstürzende Excel-Mappen
        blatt = nil
        ["big", "alt", 2, 1].each do |index|
          begin
            blatt = mappe.Worksheets(index)
          rescue
            next
          end
          trc_info :calc_workaround_blattindex=, index 
          break
        end
        blatt.UsedRange.Calculate if blatt
      rescue
        trc_aktuellen_error "Beim MappenAbsturz-Workaround für ExcelBerecnung", 5
      end   
      excel.CalculateFull # #?# nochmal rescue?
    end
  end
  
end

FARBEN = {
:eaExakt => 35, #'RGB(192, 255, 128) #' BlassGrün
:eaExaktMitRundung => 35, #'RGB(192, 255, 128) #' BlassGrün
#'              FARBEN[:eaExaktMitRundung => RGB(0, 255, 0)      #' "Grelles Grün"
#'              FARBEN[:eaExaktMitRundung => RGB(255, 255, 128)  #' BlassGelb
:eaFastExakt => 4,       #' => RGB(128, 255, 0) #' "Gelbgrün"
#'               => RGB(255, 255, 128) #' "Hellgelb"
#'              => RGB(192, 255, 128)  #'BlassGrün
#'               => RGB(0, 255, 0) #' "Grelles Grün"
#'               => RGB(255, 255, 0) #'Gelb
:eaNochOK => 6,        #'=> RGB(255, 255, 0) #'Gelb
#'               => RGB(255, 255, 128) #' "Hellgelb"
#'               => RGB(255, 192, 0) #'offiziell: "Gold"
#'               => RGB(255, 192, 128) #'blasses Rot, offiziell: "Gelbbraun"
:eaUngenau => 44,      #'=> RGB(255, 192, 0) #'offiziell: "Gold"
#'               => RGB(255, 128, 128) #'helleres Rot, gibt#'s nicht in der Palette
#'               => RGB(255, 192, 192) #'offiziell: "Hellrosa", aber Touch ins violette
#'               => RGB(255, 192, 128) #'blasses Rot, offiziell: "Gelbbraun"
#'               => RGB(255, 255, 128) #'BlassGelb
#'               => RGB(255, 255, 0) #'Gelb
:eaZuFalsch => 3,       #' => RGB(255, 0, 0) #'"Rot"
#'               => RGB(255, 0, 255) #'offiziell: "Rosa", (auch Touch ins violette)
#'               => RGB(255, 192, 192) #'offiziell: "Hellrosa", aber Touch ins violette
#'             => RGB(255, 128, 255) #' das ergibt die gleiche Farbe (Rosa-Violett)
#'               => RGB(255, 192, 128) #'blasses Rot, offiziell: "Gelbbraun"
:eaXceptioon => 39,          #' => RGB(192, 128, 255) #' offiziell "Lavendel"
#'            .Interior.Color => RGB(128, 192, 255) #' BlassBlau
:eaKeinErgebnis => 15 #' => RGB(192, 192, 192) #' Hellgrau ("Grau-25%")
}

#FARBEN.each { |te, farbe| FARBEN[te]=farbe+16 }

class ExstarError < ExcelError  
end





end # if not defined? ...




if __FILE__ == $0 then

  durchlaufe_unittests($0)

end


__END__

Code für Vergleich ohne Visual Basic:

=begin
#ok  RefSheetAdresse = "'[" & RefMappenName & "]" & Blatt.Name & "'!"
    letzte_zelle = blatt.Cells.SpecialCells(XLCellTypeLastCell)
  #Set LetzteZelle = Blatt.Cells.SpecialCells(xlCellTypeLastCell)

#    (1..letzte_zelle.Row).each { |z_nr|
 #     (1..letzte_zelle.Column).each { |sp_nr|
      erste_zelle = blatt.Cells.Find("what"=>"=", "LookIn"=>-4123, "LookAt"=>2, "SearchOrder"=>1, "MatchCase"=>true)
      z_1 = erste_zelle.Row
      sp_1 = erste_zelle.Column
      zelle = erste_zelle
      loop {
        #trace "Z,s", [z_nr , sp_nr]
        #zelle    =    blatt.Cells(z_nr, sp_nr)
        #trace :z, [zelle.row, zelle.column]
        refzelle = refblatt.Cells(zelle.row, zelle.column)

        #' #*# diese Varianten über Parameter nach auï¿½en fï¿½hren:
  #'      If Zelle.Interior.ColorIndex > 0 Then
        if refzelle.Interior.ColorIndex > 0
  #'      If True Then
          case zelle.Interior.ColorIndex
            when 42, #' Aquamarin
            8,  #' Türkis
            #'when 6,  #' gelb
            10, #' Grün  #' veraltet
            16, #' mittelgrau (50%)
            48, #' mittelhellgrau (40%)
            15 #' hellgrau (25%)
              #' nix tun
            else
      #'        If Range(RefSheetAdresse & Zelle.Address(False, False)).Interior.ColorIndex > 0 Then
             # If True Then
                begin
                  vgl_erg = zelleVerglKommentieren(zelle, refzelle, false)
                  if vgl_erg != :eaKeinErgebnis
                  #' damit dieser Wert frei wird, um Errors zu zählen
                    zaehle_erg(vgl_erg)
                  end
                rescue
                  trace :vglrescue, $!
                  zaehle_erg :eaKeinErgebnis
                end
              #End If
          end #case
        end  #If RefZelle.Interior.ColorIndex > 0 Then
        zelle = blatt.Cells.FindNext(zelle)
        break if zelle.Row == z_1 and zelle.Column == sp_1
      }
    #}
  }
=end


def zaehle_erg(erg)
  $zaehler[erg] += 1
  #trace $zaehler
end

def zelleVerglKommentieren(zelle, refzelle, formelnKorrigieren = false)
  #trace :zelle, zelle.Value
  #trace :z_value, zelle.ole_get_methods
  zellvalue = zelle.value

  return :eaKeinErgebnis if zellvalue.nil? #empty?

  #return :eaKeinErgebnis if zelle.Formula !~ /^=/
    # and If VarType(Zelle) = vbString Then Exit Function


    refvalue = refzelle.Value
    return vglerg = teste_zelle(zellvalue, refvalue) #, formelnKorrigieren)

    #zelle.Activate
    zelle.Interior.ColorIndex = FARBEN[vglerg]

    #trace :vglerg, vglerg
    if vglerg != :eaExakt
      #' Dann brauchen wir in jedem Falle einen Kommentar
      if zelle.Comment.nil?
        zelle.AddComment
        zelle.Comment.Shape.Width = 180
        zelle.Comment.Shape.Height = 120
        kommentieren = true
      else
        kommentieren = zelle.Comment.Text =~ /^Test/i
          ###GoTo NoComment #' Anderer Kommentar - Dann tun wir nichts
      end
    end #' nicht exakt

    if kommentieren
      comment = zelle.Comment
      case vglerg
      when :eaExakt
        if zelle.Comment
          #comment = zelle.Comment
          if comment.Text =~ /^Test/i
            if comment.Text =~ /.+OK/
              comment.Delete
            else
              comment.Text "Test vom " + Time.now.strftime("%d.%m. %H:%M") + " OK"
            end
          end
        end
        zelle.Interior.ColorIndex = FARBEN[:eaExakt]  #'BlassGrün
      when :eaXceptioon
        zelle.Comment.Text "Test vom " + Time.now.strftime("%d.%m. %H:%M")  +
                      "\nmit Fehler Nr=" #& Err.Number
                      "\n''" + $!.to_s + "''" +
                      "\nErg=" + zellvalue.to_s +
                      "\nRef=" + refvalue.to_s
        zelle.Interior.ColorIndex = FARBEN[:eaXceptioon]  #' offiziell "Lavendel"
      else
        kommentarText = "Test vom " + Time.now.strftime("%d.%m. %H:%M")  +
                 "\n\nErg= " + zellvalue.to_s +
                 "\nRef= " + refvalue.to_s
        if zellvalue.is_a?(Numeric) and refvalue.is_a?(Numeric)
          diff = zellvalue - refvalue
          if diff.abs > 0.009
            kommentarText +=
                 "\n\nDiff= " + diff  #Round(diff, 12 + 2 - Len(Int(100 * RefValue)))
                 #' rundet die Differenz auf 12 signifikante (bezogen auf den Ref-Wert) Stellen
                 #' (außer wenn Abs(RefValue) < 0.01 (also fast Null) ist, dann wird einfach auf 13 Nachkommastellen gerundet)
            if vglerg != :eaExakt and vglerg != :eaFastExakt and vglerg != :eaExaktMitRundung
              relAbw = 42 #' ist groß genug, falls der folgende Zweig nicht genommen wird
              if refvalue != 0
                relAbw = diff / refvalue.abs
                kommentarText +=
                   "\nRelAbw= " + "%.8f" % relAbw #Format(RelAbw, "0.000000000%")
              end
              if zellvalue != 0 and relAbw.abs > 0.001
                relAbw = diff / zellvalue.abs
                kommentarText += " (/Ref)" +
                   "\nRelAbw= " + "%.8f (/Erg)" % relAbw #Format(RelAbw, "0.000000000%")
              end
            end
          end #'Abs(Diff) > 0.009
        end #' IsNumeric
        zelle.Comment.Text kommentarText
            #'" (" & Adresse & ")"
      end # case else
    end # if kommentieren
    if false
    case vglerg
      when :eaFastExakt
        zelle.Interior.ColorIndex = FARBEN[:eaFastExakt]  #' "Gelbgrün"
      when :eaNochOK
        zelle.Interior.ColorIndex = FARBEN[:eaNochOK] #'Gelb
        #'.Interior.ColorIndex = &H6000
      when :eaUngenau
        zelle.Interior.ColorIndex = FARBEN[:eaUngenau] #'offiziell: "Gold"
      when :eaZuFalsch
        zelle.Interior.ColorIndex = FARBEN[:eaZuFalsch] #'"Rot"
      else
        trc_info :else
        zelle.Interior.ColorIndex = FARBEN[:eaKeinErgebnis] #' Hellgrau
        #' VergleichsErgebnis weder Ok noch Error
    end #'case VergleichsErgebnis
    end
    vglerg
=begin
    pos = InStr(.Formula, "+")
    If pos > 0 Then
      If Mid$(.Formula, pos + 1, 2) <> "0+" Then _
        pos = 0
    End If

    If pos = 0 Then
      .Interior.Pattern = xlPatternAutomatic
      .Interior.PatternColorIndex = xlColorIndexAutomatic
    Else
      .Interior.Pattern = xlPatternLightUp
'            .Interior.PatternColor = RGB(64, 255, 32) '"Grelles Grün"
'            .Interior.PatternColor = RGB(64, 255, 96) '"Meeresgrün"
      .Interior.PatternColor = RGB(64, 128, 16) '"Grün"
    End If
  End With 'Zelle
  ZelleVerglKommentieren = VergleichsErgebnis
End Function
=end
end

def teste_zelle(zellvalue, refvalue) #, korrigieren = false)

#    refvalue  = refzelle.Value
    return :eaXceptioon if zellvalue == -2146826273
    #trace :z, zellvalue
    refvalue  += 0.7 if refvalue.is_a?(Numeric)
  #  zellvalue = zelle.value

   #   formel = zelle.formula
    #  hatplus = formel =~ /\+0\+/
#      plusPos = InStr(.Formula, "+")
#ok      If PlusPos > 0 Then HatPlus = Mid$(.Formula, PlusPos + 1, 2) = "0+"
=begin
      #' Nicht Korrigieren, dann entfernen wir ein eventuelles Plus
      if ! korrigieren and hatplus
          #"=" & Mid$(.FormulaLocal, PlusPos + 3)
        zelle.FormulaLocal = formel.sub(/=[\d.e+-]+\+0\+/, '=')
        hatplus = false
      end


      pluswert =  if hatplus
                    formel.match(/=([\d.e+-]+)\+0\+/)[1]
                  else 0 end
        #PlusWert = CDbl(Mid$(.FormulaLocal, 2, PlusPos - 1 - 1))
=end
      #' CurrentArray: Neu für Excel 2002, damit dort auch Bereiche funktionieren
      #' Dann werden manchmal Berechnungen überflï¿½ssigerweise mehrfach angestoï¿½en
      #' vielleicht Calculate komplett weglassen? #?#
      if false ##JedeZelleVorVergleichenNeuberechnen Then
=begin
        If .HasArray Then
          .CurrentArray.FormulaArray = .CurrentArray.FormulaArray
          .CurrentArray.Calculate
        Else
          .Formula = .Formula  #'###
          .Calculate
        End If
=end
      end

      if zellvalue.is_a?(String) # == VT_STRING
      #trace :vt, varType(zelle.value)
        erg = zellvalue
        if erg == refvalue
          return :eaExakt
        else
          return :eaZuFalsch
        end
      end

      erg = zellvalue #- pluswert
      real_erg = erg.to_f
#    End With #'Zelle
    differenz = real_erg - refvalue
    if differenz == 0
      return :eaExakt
#    end
    else
      #'#*# Gruppieren nach Vergleichsart
      absdiff = differenz.abs
      absDiffRelativ = absdiff * 2 / (real_erg.abs + refvalue.abs)

      if absDiffRelativ < 0.000000000000004 then #'eE-15
        return :eaExakt
        #' Das war die zweite und letzte Chance für teExakt
      elsif absdiff < 0.0051
        return :eaExakt #MitRundung
      elsif absDiffRelativ < 0.0000000002 then #'2e-10
        return :eaExakt #MitRundung
      else
=begin
        if not hatplus
          if zelle.Text =~ /\#\#$/ then
            gerundetesErg = (real_erg*100).round / 100.0
          else
            gerundetesErg = zelle.Text.to_f
          end
        else
=end
          gerundetesErg = (real_erg*100).round / 100.0 #Round(RealErgWert, 2)
#        end
        if gerundetesErg.to_s ==( ".2f" % refvalue)
          return :eaExakt #MitRundung
          #' Das war die letzte Chance für teExaktMitRundung
        elsif (gerundetesErg - refvalue).abs < 0.01001
          return :eaFastExakt
        elsif absDiffRelativ < 0.000001 then #'1E-6
          return :eaFastExakt
          #' erlaubt bei 1 000 000 +-10, bei 250 000 +-0.25 (=LogFakt) [, bei 10 000 +-0.01]
        else
          if refvalue.abs < 500
            logFaktor = 0
          else
            logFaktor = Math::log(refvalue) / Math::log(10) - 1 #' bei 100->1, bei 500->1.7 bei 1E4->3
            logFaktor = logFaktor * logFaktor * logFaktor #' bei 1000->8, bei 10000->27
          end
          if absdiff < 0.002 * logFaktor then #' bei 500 wird die 0.01-Grenze überschritten
            #' bei 1000->0.016, bei 1E4->0.054, bei 1E5->0.13, bei 1E6->0.25, [bei 1E7->0.43, 1e8->0.68]
            #' also:  1.6E-5           5.4E-6          1.3E-6       [2.5E-7         4.3E-8         6.8E-9] < 1E-6=AbsRelativSchranke
            return :eaFastExakt
            #' Das war die letzte Chance für teFastExakt
          elsif absdiff < 0.05
            return :eaNochOK
          elsif absDiffRelativ < 0.00025 #'2.5E-4
            #' erlaubt bei 100000 +-25, bei 8000 +-2 (=QuadDiff) [bei 200 +-0.05 (=AbsDiff)]
            return :eaNochOK
          elsif erg == erg.to_i and absdiff == 1 then #'Bei Rundung auf ganze Euros
            return :eaNochOK
          else
            relQuadDiff = absdiff * absDiffRelativ
            if relQuadDiff < 0.0005 then #'5E-4
              #' erlaubt [bei 200000 +-10,] bei 8000 +-2 (=RelDiff), bei 200 +-0,1, bei 5 +- 0.05 (=AbsDiff)
              return :eaNochOK
              #' Das war die letzte Chance für teNochOK
            elsif absdiff < 0.22 then #' Das kann aber z.B. auch bedeuten Erg=0.1, Ref=-0.1
              return :eaUngenau
            elsif absDiffRelativ < 0.01 then #'1E-2
              #' erlaubt bei 1000 +-10, bei 250 +-2.5 (=QuadDiff) [, bei 100 +-1, bei 22 +-0.22 (=AbsDiff) [bei 10 +-0,1]]
              return :eaUngenau
            elsif relQuadDiff < 0.025 #'2.5E-2
              #' erlaubt [bei 4000 +-10,] bei 250 +-2.5 (=RelDiff),  bei 40 +-1, bei 2 +-0,22 (=AbsDiff) (irrelevant: bei 1 +-0,1))
              return :eaUngenau
              #' Das war die letzte Chance für teUngenau
            else #' bleibt nur noch:
              return :eaZuFalsch
            end # if' QuadDiff
          end #' AbsLogVgl
        end #' GerundetesErg
      end #' schwierigster Fall
    end #' Differenz <> 0

=begin
    If return :eaExakt Or return :eaExaktMitRundung Then   #'teExakt
      If HatPlus Then
        With Zelle
          .Formula = "=" & Mid$(.Formula, PlusPos + 3)
        End With #'Zelle
      End If
    Else
      If Korrigieren Then
        With Zelle
        Dim AlteFormelanfangsPos  As Integer
          If Left$(.FormulaLocal, 1) <> "=" Then
            AlteFormelanfangsPos = 1
          Else
            AlteFormelanfangsPos = IIf(HatPlus, PlusPos + 3, 2)
          End If
          If Abs(-Differenz - PlusWert) > (Abs(Differenz) + Abs(PlusWert)) * 0.000000000000002 Then
            .FormulaLocal = "=" & Format(-Differenz, "0.0#####E-0##") & "+0+" _
                            & Mid$(.FormulaLocal, AlteFormelanfangsPos)
          End If
        End With #'Zelle
      End If
    End If #' TesteZelle nicht OK

  On Error GoTo 0

  Exit Function

Fehler:
  return :eaXceptioon


=end
end


