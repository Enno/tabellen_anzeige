
if not defined? HauptKonfigHash then

require 'schmiedebasis'
require 'models/system_fkt/systemebene'

class Symbol
  def <=>(other)
    self.to_s <=> other.to_s
  end
end


require 'delegate'

class Zeit < DelegateClass(Fixnum)
#  undef :==
  def initialize(i)
    if i.respond_to?(:=~) and i =~ /^\s*\d+\s*\|\s*\d+/ then
      j,m = i.split("|").map{|s| s.to_i}
      i = j * 12 + m
    end
    super(i.to_i)
  end

  def self.jm(jahr, monat)
    new(jahr * 12 + monat)
  end

  def j
    self / 12
  end

  def m
    self % 12
  end

  def +(x)
    Zeit.new(to_i + x)
  end
  def -(x)
    Zeit.new(to_i - x)
  end

  def eql?(x)
    x.is_a?(Zeit) and to_i == x
  end

  def to_s
    if self.m == 0 and self.j > 1000 then
      "#{j-1}|12"
    else
      "#{j}|#{m}"
    end
  end

  def inspect
    "<zeit:#{to_s}>"
  end

  def hash
    to_i.hash*65536
  end
=begin
  def to_i
    @wert
  end
  def to_s
    to_i.to_s
  end
=end
end



class VergleichsErgebnis < Hash
  ERGEBNIS_ARTEN = [ :eaAlles,
                        :eaExakt,
                        #:eaExaktMitRundung,
                        :eaFastExakt,
                        :eaNochOK,
                        :eaUngenau,
                        :eaVollDaneben,
                        :eaXception]

  AA_INFO = {
    :eaAlles        => {:kurz=>"A"},
    :eaExakt        => {:kurz=>"E"},
    #:eaExaktMitRundung,
    :eaFastExakt    => {:kurz=>"F"},
    :eaNochOK       => {:kurz=>"N"},
    :eaUngenau      => {:kurz=>"U"},
    :eaVollDaneben  => {:kurz=>"V"},
    :eaXception        => {:kurz=>"X"}
  }

  def initialize(ini_arg = nil, vglzeit = nil)
    self.default = 0
    ini_arg = self.class.str_zu_hash(ini_arg) unless ini_arg.is_a?(Hash)
    @zeit = nil # erstmal initialisieren (keine Warnungen mehr)
    self.update ini_arg
    @zeit = OrtPersistenz.zeit2str(vglzeit) if vglzeit
  end

  def self.str_zu_hash(ini_str)
    erg_hash = Hash.new(0)
    return erg_hash if ini_str.nil?

    zugewiesen = false
    ERGEBNIS_ARTEN.each {|ea|
      ea_kurz = AA_INFO[ea][:kurz]
      #trace :str_zu_hash, [ini_str, ea_kurz, ini_str =~ /#{ea_kurz}: ?(\d+)/]
      if ini_str =~ /#{ea_kurz}: ?(\d+)/
        erg_hash[ea] = $1.to_i
        zugewiesen = true
      end
    }
    if ini_str =~ /Zeit: ?([^,]+),/ then
      erg_hash[:zeit] = $1
      zugewiesen = true
    end
    if ini_str =~ /Bib: ?(.+)$/ then
      erg_hash[:bib] = $1
      zugewiesen = true
    end
    if zugewiesen then
      #trace :erg_hash, erg_hash
      erg_hash
    else
      raise "Vergleichsergebnis kann aus diesem Wert nicht gelesen werden (wert=#{ini_str.inspect})"
    end
  end

  def to_s
    "Zeit:#{@zeit}, " +
    ERGEBNIS_ARTEN.map { |abw_art|
      AA_INFO[abw_art][:kurz] + ":#{self[abw_art]}"
    }.join(", ") +
    ", Bib:#{bib}"
  end

  def [](*args)
    args.inject(0) { |sum, erg_art|
      sum + super(erg_art)
    }
  end

  def []=(art, anzahl)
    super(art, anzahl)
    raise "nil auf VglErg zugewiesen" unless anzahl
  end

  def ==(anderes_vgl_erg)
    return false unless self.class === anderes_vgl_erg
    ERGEBNIS_ARTEN.select {|ea|
      self[ea] != anderes_vgl_erg[ea]
    }.empty?
  end

  attr_reader :zeit

  def zeit=(neue_zeit)
    @zeit = OrtPersistenz.zeit2str(neue_zeit)
  end

  def update(anderes)
    kop = anderes.dup
    if anderes.has_key?(:zeit) then
      self.zeit = anderes[:zeit]
      kop.delete(:zeit)
    end
    if anderes.has_key?(:bib) then
      self.bib = anderes[:bib]
      kop.delete(:bib)
    end
    super(kop)
  end

  def bib
    @bib ||= ""
  end

  def bib=(bibname)
    @bib = bibname
  end

  def summe
    return self[:eaAlles]
    #return 0 if not self.exist?(:vgl)
    #self.values.inject(0) {|sum, w|  sum + w }
  end

  def anz_unbekannt
    self.values.inject(summe*2) {|noch_unbek, anz|  noch_unbek - anz }
  end

  def +(summand)
    #trace :plus, to_s
    if not summand.is_a?(VergleichsErgebnis) then
      raise TypeError, "#{summand.inspect} ist kein VergleichsErgebnis"
    end
    erg = self.class.new({}, [self.zeit, summand.zeit].compact.min)
    if summand
      ERGEBNIS_ARTEN.each { |art|
        erg[art] = self[art] + summand[art]
      }
      erg.bib = ( (bib and bib > "") ? bib : summand.bib )
    else
      erg = summand.dup
    end
    erg
  end

  def ampelwert
     abweichungen = self.reject {|abw_art, anz|
        anz == 0 or
        [:eaAlles, :eaExakt, :eaFastExakt].include?(abw_art)
    }
    if summe > 0 and
       ( abweichungen.empty? or
         (abweichungen.keys == [:eaNochOK] and self[:eaNochOK] <= 0.01 * summe)
       ) then
      :gut
    elsif self[:eaXception, :eaVollDaneben] <= 0.01 * summe
      :naja
    else
      :schlecht
    end
  end

end


class EintragPersistenz
  def initialize()
    @erg_vertr = {}
    @erg_werte = {}
  end

  attr_accessor :kurzinfo, :beschreibung, :quellort, :zielort

  def vglerg_neu_werte(vglerg)
    @erg_werte[OrtPersistenz.zeit2str(vglerg.zeit)] = vglerg
  end

  def vglerg_neu_vertr(vglerg)
    @erg_vertr[OrtPersistenz.zeit2str(vglerg.zeit)] = vglerg
  end

  def vglerg_liste_werte
    @erg_werte.values
  end

  def vglerg_liste_vertr
    @erg_werte.values
  end

  def vglerg_werte(zeit)
    @erg_werte[OrtPersistenz.zeit2str(zeit)]
  end

  def vglerg_vertr(zeit)
    @erg_vertr[OrtPersistenz.zeit2str(zeit)]
  end

  def []=(inifile_schluessel, wert)
    #trc_temp "inifile_schluessel, wert", [inifile_schluessel, wert]
    if inifile_schluessel.to_s =~ /^(erg_(vertr|werte)_)?(20\d\d-\d\d-\d\d.*)$/
      prefix, art, zeit = $1, $2, $3
      vglerg = VergleichsErgebnis.new(wert, zeit)
      #trc_temp :vglerg_gelesen, vglerg
      if art != "vertr" then
        @erg_werte[zeit] = vglerg
      else
        @erg_vertr[zeit] = vglerg
      end
    else
      instance_variable_set("@#{inifile_schluessel}", wert)
    end
  end


  def sort
    self.instance_variables.inject([]) do |paare, inst_var|
      iv_name = inst_var[1..-1].to_sym  # ohne @
      iv_wert = instance_variable_get(inst_var)
      paare + if inst_var =~ /^@(erg_.....)/ then
        iv_wert.map do |zeit, vglerg|
          ["#{iv_name}_#{zeit}".to_sym, vglerg]
        end
      else
        [[iv_name, iv_wert]]
      end
    end
  end

  # zum UnitTest
  def == (anderes)
    not self.instance_variables.find do |iv|
      if iv =~ /^@erg_(.....)/ then
        m = "vglerg_liste_#{$1}".to_sym
        self.send(m).sort != anderes.send(m).sort
      else
        iv_sym = iv[1..-1].to_sym  # ohne @
        instance_variable_get(iv) != anderes.send(iv_sym)
      end
    end
  end

end

class OrtPersistenz

  @@PfadZuSerie = {}

  attr_reader :dateiname

  def initialize(pfad) # pfad ist anzeige_pfad, auï¿½er bei ExlOrd, da wird auch Excel-Dateien genmommen
    @pfad = pfad.chomp("/")
    realpfad, kern_name = *if @pfad =~ /^(.*)\/<(.*)>/ then
      [$1, $2]
    else
      [@pfad, "OrdnerInfo"]
    end
    @dateiname = realpfad + "/TS_#{kern_name}.ats"

    # #*# Das folgende ist nur für die Zeit des ï¿½bergangs:
    if not File.exist?(@dateiname) and File.exist?(realpfad + "/NTestserien-Ergebnisse.txt") then
      File.rename(realpfad + "/NTestserien-Ergebnisse.txt", @dateiname)
    end

#    if pfad =~ /Excel-Dateien\/?$/ then
 #     not File.exist?(@dateiname) and File.exist?(realpfad + "/TS_OrdnerInfo.ats") then
  #    File.rename(realpfad + "/TS_OrdnerInfo.ats", @dateiname)
   # end


    einlesen #if File.exist?(@dateiname)

  end

  def self.fuer_pfad(pfad)
    trc_info :fuer_pfad, pfad
    @@PfadZuSerie.fetch(pfad) {
      @@PfadZuSerie[pfad] = new(pfad)
    }
    #trace :fuer_pfad_hash, @@PfadZuSerie[pfad]
    @@PfadZuSerie[pfad]
  end

  def loeschen
    @hash.keys.each { |eintr_name|
      @hash.delete(eintr_name)
    }
  end

  def self.zeit2str(zeit)
    if zeit.respond_to?(:strftime) then
      zeit.strftime("%Y-%m-%d_%H%M")
    else
      zeit
    end
  end

  def neues_werte_erg(dateiname, erg)
    trc_info "serg neues_erg datei, erg", [dateiname, erg.to_s]
    inhalt_fuer_datei(dateiname)["werte_#{self.class.zeit2str(erg.zeit)}".to_sym] = erg
  end

  def inhalt_fuer_eintrag(dateiname)
    #return nil unless dateiname
    eintragname = File.basename(dateiname)
    #trc_temp "inhalt_fuer_datei", eintragname
    @hash.fetch(eintragname){
      @hash[eintragname] = EintragPersistenz.new # #*# DRY machen
    }
  end

  def einlesen
    trc_info :dateiname, @dateiname
    @hash = InifileHash.new(@dateiname) do |neuer_abschnittsname|
       trc_temp :neuer_abschnittsname, neuer_abschnittsname
       if neuer_abschnittsname == "*META*" then
         Hash
       else
         EintragPersistenz
       end.new
    end
    # #*# diese Information soll in der Zukunft beim Einlesen ï¿½berprï¿½ft werden (auï¿½er pfad)
    @hash["*META*"].update({:dateityp => "ort-info",
                            :version => "0.1",
                            :pfad => @pfad})
#    @rohash
 #   @rohash.each do |eintragname, inhalt|
  #    @hash[eintragname] = EintragPersistenz(inhalt)
   # end
  end

  def speichern
    #@hash.each do
    @hash.flush
  end

  def alle_zeiten
    @hash.values.inject([]) { |zeiten_bisher, zeit_erg_hash|
      zeiten_bisher + zeit_erg_hash.keys
    }.uniq
  end

end



#########################


class Option < Struct.new(:symbol, :default, :text, :validierung, :wirkung, :wert, :historie)
  def initialize(*args)
    super
    self.speicherwert = default if self.wert.nil?
  end
  
  def wert=(neu_wert)
    #trc_temp "neu_wert", neu_wert
    neu_wert = default if neu_wert.nil?
    super(neu_wert)
    self.wirkung.call(neu_wert) if self.wirkung
    #trc_temp "wert_gesetzt", wert
  end

  def historie=(neu_historie)
    #trc_temp :neu_hist, neu_historie
    neu_historie = default if neu_historie.nil? or neu_historie == ""
    super(neu_historie)
    return unless historie
    neu_wert = historie.split("^")[0]
    if neu_wert == "" then
      neu_wert = default
      super(neu_wert+neu_historie)
    end
    self.wert = neu_wert
  end

  def speicherwert
    default.is_a?(String) ? historie : wert
  end

  def speicherwert=(neu)
    #trc_temp :spwert_neu, neu
    case default
    when Symbol then
      self.wert = neu.to_sym if neu and neu != ""
    when String then
      self.historie = neu
    else
      self.wert = neu
    end
  end

end

class Opts

  def initialize(neuhash)
    @intern = {}
    OPTIONEN.each do |abschn_sym, elemente|
      elemente.each { |defaultwert, symbol, anzeige_text, validierung, wirkung|
        neu_opt = Option.new(symbol, defaultwert, anzeige_text, validierung, wirkung)
 #       trc_temp :neuhash, neuhash[symbol]
        neu_opt.speicherwert = neuhash[symbol]
        trc_temp :neu_opt, neu_opt
        @intern[symbol] = neu_opt
      }
    end
  end

  def [](schluessel)
    intern_sicher(schluessel).wert
  end

  def speicherwert(schluessel)
    intern_sicher(schluessel).speicherwert
  end

  def neuer_speicherwert(schluessel, neu)
    intern_sicher(schluessel).speicherwert = neu
  end

  def speicherwerte
    erg = {}
    @intern.each do |sym, opt|
      erg[sym] = opt.speicherwert
      trc_temp sym, opt.speicherwert
    end
    erg
  end

private 
  def intern_sicher(schluessel)
    @intern[schluessel] || raise("Bug: #{schluessel.inspect} ist kein Opts-Schlüssel")
  end
  
end


class HauptKonfigHash
  attr_reader :start_orte, :allg, :opts
  def initialize(dateiname)
    trc_temp :hauptkonfig_init_start, dateiname
    @inihash = InifileHash.new(dateiname)
    @start_orte = @inihash.dup
    @start_orte.delete("Allgemein")
    @start_orte.delete("Optionen")
    @allg = @inihash["Allgemein"] #, {})
    @opts = Opts.new(@inihash["Optionen"])
    trc_temp :hauptkonfig_init_fertig
  end

  def save
    @inihash = @start_orte.dup
    @inihash.each{ |name, inhalt|
      inhalt.delete(:place)
    }
    @inihash["Allgemein"] = @allg
    @inihash["Optionen"]  = @opts.speicherwerte
    @inihash.flush
  end

end


class WandelArt < Struct.new(:bez, :vorlagen_name, :vstruktur)
  attr_reader :vorlagenpfad_sym
  def initialize(*args)
    super
    @vorlagenpfad_sym = ("ExstarVorl_" + self.bez).to_sym
  end
end

WA_NG1 = WandelArt.new "NG1", "E2Vorlage.xls",  :vmappe
WA_HR5 = WandelArt.new "HR5", "E2VorlHR5a.xls", :vmappe
WA_HR6 = WandelArt.new "HR6", "E2VorlHR6.xls",  :vzeilen

WANDEL_ARTEN = [WA_NG1, WA_HR5, WA_HR6]

WANDEL_ARTEN_STR_ZU_OBJ = {}
WANDEL_ARTEN.each do |wa_obj|
  WANDEL_ARTEN_STR_ZU_OBJ[wa_obj.bez] = wa_obj
end

OPTIONEN = {}

WANDEL_ARTEN.each do |wa_obj|
  opt_sym = ("Excel_Vorlage_" + wa_obj.bez).to_sym 
  OPTIONEN[opt_sym] = [
    [
      KONFIG_DIRNAME + "/" + wa_obj.vorlagen_name,
      wa_obj.vorlagenpfad_sym,
      "Dateiname der Vorlage",
      proc do |vorlage_dateiname| 
        "Problem: Die Exstar-Vorlage-Datei existiert nicht!" if not File.exist?(vorlage_dateiname)         
      end  
    ],[
      "#{WORK_DIRNAME}/Exstar#{wa_obj.bez}",
      ("Dbf2ExlAusgabePfad_" + wa_obj.bez).to_sym,
      case wa_obj.vstruktur
      when :vmappe  then "Basis-Ordner für die erzeugten Excel-Dateien:"
      when :vzeilen then "Ziel-Dateiname für die erzeugten Excel-Daten:"
      else
        trc_info "falsche vstruktur", [wa_obj.vstruktur, wa_obj]
        "Ort für die erzeugten Excel-Dateien:"
        nil
      end
    ]       
  ]
end

OPTIONEN.update( {
  :Dbf_zu_Excel_Wandlung => [
    [:raise, :WennBeiExcelErstellungMappeExistiert,
            ["Falls Excel-Testfälle schon existieren",
              [:overwrite, "ohne Nachfrage ï¿½berschreiben"],
              [:raise,     "Fehler melden"],
              [:excel,     "Excel nachfragen lassen"]
            ]],
    [true,  :Dbf2ExlLeerzeichenInVsnrErsetzen, "Leerzeichen in Versicherungsnummern mit Nullen ersetzen"],
    ["M:/MathStar/DB/GENERALI/DIV",
            :MathstarDivOrdner, "Ordner für Mathstars Dividenden-Dateien"]
  ],
  :Excel_Vergleichsoptionen => [
    [true,  :NurZellenMitFormelnVergleichen, "Nur Zellen mit Formeln vergleichen"],
    [true,  :BerechnenVorVgl,             "Vor dem Vergleichen berechnen"],
    [true,  :FeldFarbBeimVgl,             "Felder fürben entsprechend ï¿½bereinstimmung"],
    [true,  :MappeSpeichernBeimVgl,       "Mappen nach dem Vergleichen speichern"],
    [false, :ExcelImmerWiederNeuStarten,  "Excel vor jedem Vergleich neu starten"]
  ],
  :Prozess_Visualisierung => [
    [true, :OrtsAnzeigeSynchronisieren,   "Orts-Anzeige &synchron zum gerade durchlaufenen Ort wechseln."],
    [true, :MappenAnzeigenBeiBatchjobs,        "Excel-Mappen wï¿½hrend automatisierter Prozesse &anzeigen"]
  ],
  :Exstar_Bibliothek => [
    ["",
      :VorgegebeneBibliothek,       
      "&Vorgegebene Kern-Bibliothek:",
      proc do |bib_dateiname|
        if KONFIG.opts[:VorgBibliothekNutzen] then
	        if bib_dateiname == "" then
	          "Problem: Es wurde noch keine Exstar-Bibliothek angegeben!"
	        elsif not File.exist?(bib_dateiname) then
	          "Problem: Vorgegebene Exstar-Bibliothek existiert nicht!"
	        end 
        end
      end
    ],
    [true, :VorgBibliothekNutzen,        "Vorgegebene Kern-Bibliothek &benutzen"]
  ],
  :Testschmiede_Optionen => [
#    [ !RELEASE_FFM, 
 #     :Expertenmodus,   
  #    "E&xperten-Features aktivieren",
   #   proc
    #],
    [ $trace_stufe, 
      :TraceStufeHaupt,    
      ["Grad der &Protokollierung des Hauptprogramms",
        [:fehler,  "Nur bei &Fehlern"],
        [:hinweis, "Bei Fehlern und &Hinweisen"],
        [:info,    "&Ausfï¿½hrlich, ohne Debugging-Details"],
        [:temp,    "&Komplett (nur zum Debuggen)"]
      ],
      nil,
      proc {|neu_wert| $trace_stufe = neu_wert }
      
    ],
    [:info,    :TraceStufeDiener,      ["Grad der Protokollierung der Dienste",
                                          [:fehler,  "Nur bei F&ehlern"],
                                          [:hinweis, "Bei Fehlern und H&inweisen"],
                                          [:info,    "A&usfï¿½hrlich, ohne Debugging-Details"],
                                          [:temp,    "K&omplett (nur zum Debuggen)"]
                                        ]
    ]
  ]
})

GRUPPEN = {
  :Dbf_zu_Excel_NG1 => [
    :Dbf_zu_Excel_Wandlung, :Excel_Vorlage_NG1, :Exstar_Bibliothek  #, :Prozess_Visualisierung
  ],
  :Dbf_zu_Excel_HR5 => [
    :Dbf_zu_Excel_Wandlung, :Excel_Vorlage_HR5 , :Exstar_Bibliothek #, :Prozess_Visualisierung
  ],
  :Dbf_zu_Excel_HR6 => [
    :Dbf_zu_Excel_Wandlung, :Excel_Vorlage_HR6 , :Exstar_Bibliothek #, :Prozess_Visualisierung
  ],
  :Dbf_Exl_Vgl => [
    :Dbf_zu_Excel_Wandlung, :Excel_Vorlage_NG1, :Exstar_Bibliothek, :Excel_Vergleichsoptionen
  ],
  :Excel_Vergleich => [
    :Excel_Vergleichsoptionen, :Exstar_Bibliothek, :Prozess_Visualisierung
  ],
  :Hauptmenu => [
    :Dbf_zu_Excel_Wandlung, :Excel_Vergleichsoptionen, 
    :Prozess_Visualisierung, 
    :Excel_Vorlage_NG1, :Excel_Vorlage_HR6, :Exstar_Bibliothek , 
    :Testschmiede_Optionen
  ],
}
pfad_zur_Hauptkonfigdatei = KONFIG_DIRNAME + "/TestSchmiede.ini"

KONFIG = HauptKonfigHash.new(pfad_zur_Hauptkonfigdatei)

end # if not defined? HauptKonfigHash then


##########################


module KonfigUndLogFunktionen

  def exstar_vorlage_auf_neusten_stand
    trc_info :esvorl_einsprung
    
    return if defined?(EXSTAR_VORLAGE_KOPIERT)
    trc_temp :esvorl_weiter
    WANDEL_ARTEN.each do |wa_obj|
      vorl_name = wa_obj.vorlagen_name  
      backup_name = vorl_name.sub(/(\.[\w]+)$/, '_backup\1')
      begin
        require 'fileutils'
        if File.exist?("#{KONFIG_DIRNAME}/#{vorl_name}") then
          trc_info :esvorl_backup, "#{KONFIG_DIRNAME}/#{backup_name}"
          FileUtils.cp "#{KONFIG_DIRNAME}/#{vorl_name}", "#{KONFIG_DIRNAME}/#{backup_name}"
        end
        trc_info :esvorl_orig, "#{ORIG_DIRNAME}/#{vorl_name}"
        FileUtils.cp "#{ORIG_DIRNAME}/#{vorl_name}", "#{KONFIG_DIRNAME}/#{vorl_name}"
      rescue
        trc_aktuellen_error "kopieren von #{vorl_name}"
      end
    end
    
    Object::const_set(:EXSTAR_VORLAGE_KOPIERT, true)
    trc_hinweis "defined?(EXSTAR_VORLAGE_KOPIERT)", defined?(EXSTAR_VORLAGE_KOPIERT)
    
    trc_temp :esvorl_ende
  end

end # KonfigUndLogFunktionen


include KonfigUndLogFunktionen



if __FILE__ == $0
  durchlaufe_unittests($0)
end

__END__

  :Excel_Vorlage_NG1 => [
    ["#{KONFIG_DIRNAME}/E2Vorlage.xls",
      :ExstarVorl_NG1,     
      "Dateiname der Vorlage",
      proc do |vorlage_dateiname| 
        "Problem: Die Exstar-Vorlage-Datei existiert nicht!" if not File.exist?(vorlage_dateiname)         
      end       
    ],
    ["#{WORK_DIRNAME}/Exstar",
            :Dbf2ExlAusgabePfad_NG1,        "Basis-Ordner für die erzeugten Excel-Dateien:"],
  ],
  :Excel_Vorlage_HR5 => [
    ["#{KONFIG_DIRNAME}/E2VorlHR5a.xls",
      :ExstarVorl_HR5,     
      "Dateiname der Vorlage",
      proc do |vorlage_dateiname| 
        "Problem: Die Exstar-Vorlage-Datei existiert nicht!" if not File.exist?(vorlage_dateiname)         
      end       
    ],
    ["#{WORK_DIRNAME}/ExstarHR5/",
            :Dbf2ExlAusgabePfad_HR5,        "Basis-Ordner für die erzeugten Excel-Dateien:"],
  ],
  :Excel_Vorlage_HR6 => [
    ["#{KONFIG_DIRNAME}/E2VorlHR6.xls",
      :ExstarVorl_HR6,     
      "Dateiname der Vorlage",
      proc do |vorlage_dateiname| 
        "Problem: Die Exstar-Vorlage-Datei existiert nicht!" if not File.exist?(vorlage_dateiname)         
      end       
    ],
    ["#{WORK_DIRNAME}/ExstarHR6/",
            :Dbf2ExlAusgabePfad_HR6,        "Ziel-Dateiname für die erzeugten Excel-Daten:"],
  ],
