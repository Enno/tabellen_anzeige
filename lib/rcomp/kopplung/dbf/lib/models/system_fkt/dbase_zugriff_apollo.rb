
if not defined? DbfDat then

require 'schmiedebasis'

require 'win32ole'


$:.unshift "C:/ProgLang/RunRuby/lib/ruby/1.8"+ "/actrecor"
trc_info :$:, $:

ar_req_prefix = ""
require ar_req_prefix + 'active_record.rb'
trc_info "$:_nach_ar", $:
require ar_req_prefix + 'composite_primary_keys'
trc_info "$:_nach_cpk", $:

trc_info "beginne dbase_zugriff"

#2007-Mai-13 Svs: aus orte.rb hierher verlagert
require 'rdb'

EIGENES_AR_CACHING = true
#EIGENES_AR_CACHING = false

BDE_DIREKT = true

if BDE_DIREKT
  trc_hinweis "dbase-Zugriff: BDE-direkt"
  require ar_req_prefix + 'active_record/connection_adapters/apollo_bde_adapter'
else
  trc_hinweis "dbase-Zugriff: über ODBC"
  require ar_req_prefix + 'active_record/connection_adapters/apollo_adapter'
  VDATEN_DSN_NAME = "MathStarDbf-Temp"
end
trc_info :$:, $:

trc_hinweis "AR geladen, lade dbase_zugriff"


module DbfDat
  def self.oeffnen(dateiname, bestand_oder_soll)
    trc_hinweis :beginne_oeffnen_dateiname=, dateiname
    profil_setzen(File.basename(dateiname), bestand_oder_soll)
    anderer_pfad = setze_dbf_pfad(File.dirname(dateiname))
    if anderer_pfad or not VObjektBasis.connected?
      datenbank_reconnect_standard
    end
    trc_info :DbfDat_oeffnen_ok

  end

  def self.schliessen
    VObjektBasis.connection.disconnect!
    VObjektBasis.remove_connection
  end

  $aktueller_dbf_ordner = nil
  $abfragen_seit_letztem_connect = 0

private
  # returns true if dbf_pfad changed
  def self.setze_dbf_pfad( dbf_ordner )

    return false if $aktueller_dbf_ordner == dbf_ordner

    if BDE_DIREKT
      $aktueller_dbf_ordner = dbf_ordner
      return true
    else

      begin
        wsh = WIN32OLE.new("WScript.Shell")
      rescue
        raise "Konnte keine Verbindung zum Windows-Srcipting_Host herstellen um auf die Registry zuzugreifen."
        return false
      end

      reged_basispfad = 'HKCU\SOFTWARE\ODBC\ODBC.INI\\'

      dbase_orig_werte = {}

      dbase_orig_reged_pfade = [reged_basispfad, 'HKLM\SOFTWARE\ODBC\ODBC.INI\\']
      dbase_orig_reged_pfade.map! { |pfad| pfad + "dBase-Dateien\\" }
      dbase_orig_reged_pfade.each { |rpfad|
        begin # Diese Werte müssen wir haben:
          dbase_orig_werte["Driver"]   = wsh.RegRead(rpfad + "Driver")
          dbase_orig_werte["DriverId"] = wsh.RegRead(rpfad + "DriverId")
        rescue WIN32OLERuntimeError
          # nicht gefunden, na dann versuchen wir halt den nächsten.
          next
        end
        begin # Falls gefunden, nehmen wir den Rest optional mit rein:
          dbase_orig_werte["FIL"]              = wsh.RegRead(rpfad + "FIL")
          dbase_orig_werte["SafeTransactions"] = wsh.RegRead(rpfad + "SafeTransactions")
          dbase_orig_werte["UID"]              = wsh.RegRead(rpfad + "UID") # weglassen?
        rescue WIN32OLERuntimeError
        end
        break # die Suche hat ein Ende!
      }
      trc_temp "setz-dbf dbase_gefu", dbase_orig_werte
      if dbase_orig_werte == {} # in der Registry nichts brauchbares gefunden
        warne "Kein dBase-Eintrag in Registry gefunden, nehme default-Werte!"
        begin
          windows_pfad = wsh.ExpandEnvironmentStrings("%SystemRoot%")
        rescue
          warne "default windows path."
          windows_pfad = 'C:\Windows'
        end
        dbase_orig_werte["Driver"] = windows_pfad + '\System32\odbcjt32.dll'
        dbase_orig_werte["DriverId"] = '215'.hex #533
      end
      trc_temp "setz-dbf dbase_genommen", dbase_orig_werte

      neue_dsn_werte = dbase_orig_werte
      neue_dsn_werte["DefaultDir"] = dbf_ordner
      neue_dsn_werte["Description"] = "DBase-Ordner für TestWerkStatt"

      neue_dsn_werte.each { |name, wert|
        wsh.RegWrite  reged_basispfad + VDATEN_DSN_NAME + '\\' + name,  wert
      }
      #Escribo en el Key "ODBC Data Sources" para poder listar el nuevo DSN en el ODBC Manager
      wsh.RegWrite reged_basispfad + 'ODBC Data Sources\\' + VDATEN_DSN_NAME , "Microsoft dBase-Treiber (*.dbf)"
      #wsh.Popup "Fertig", 5, "Set DSNs"
      wsh = nil

      datenbank_reconnect_standard

      $aktueller_dbf_ordner = dbf_ordner
      return true
    end # dbf_ordner geändert
  end


  def self.datenbank_reconnect_standard
    if BDE_DIREKT
      adapter       = "apollo_bde"
      database      = $aktueller_dbf_ordner.gsub("/","\\")
    else
      adapter       = "apollo"
      database      = VDATEN_DSN_NAME
    end
    trc_temp :vor_db_reconnect
    datenbank_reconnect(:adapter       => adapter,
                        :database      => database      )
  end

  def self.datenbank_reconnect(spezifikation = nil)
    trc_hinweis :reconnect_spezfikation=, spezifikation
    trc_temp :reconn_anf_connected?, VObjektBasis.connected?
    alte_spezifikation = VObjektBasis.remove_connection
    spezifikation ||= alte_spezifikation
    VObjektBasis.establish_connection(spezifikation)
    $abfragen_seit_letztem_connect = 0
    trc_temp :reconn_end_connected?, VObjektBasis.connected?
    trc_info :conn_fertig
  end

  def self.sql_ausfuehren(sql_string)
    $abfragen_seit_letztem_connect += 1
    datenbank_reconnect if $abfragen_seit_letztem_connect % 42 == 0
    trc_info :sql_ausf_vorher, sql_string
    erg = VObjektBasis.connection.execute(sql_string)
    trc_info :sql_ausf_fertig, sql_string
    erg
  end

  def self.sql_select(sql_string)
    $abfragen_seit_letztem_connect += 1
    datenbank_reconnect if $abfragen_seit_letztem_connect % 42 == 0
    trc_info :sqlselect_vorher, sql_string
    erg = VObjektBasis.connection.select_all(sql_string)
    trc_info :sqlselect_fertig, sql_string
    erg
  end


  # Diese Prozedur soll sowohl für den Fall funktioniern, dass
  # profil die ".dbf"-Extension trägt, als auch nicht,
  # weiterhin auch für den Fall, dass es ein reines Profl ist
  # oder die Tabellenart und S/B vorne dranhängt.
  # #*# Die letztere Unterscheidung ist im Fall Länge=5 und Länge =4
  # heikel: "vdbes" könnte das Profil "es" oder "vdbes" meinen.
  # Hier wurde die Entscheidung für die erste Variante getroffen,
  # d.h. im Zweifelsfall _wird_ _der_ _Tabellenpräfix_ _erwartet_.
  # #*# Globales Refactoring könnte diese Frage anders lösen!!!
  def self.profil_setzen(profil, bestand_oder_soll)
    trc_hinweis :setprof_profiluebergeben, profil
    if profil =~ /^<(.+)>$/ then
      profil = $1
    else
      profil.sub!(/^(st|gz|rk|zs|v[dpktvbfa])[BS]/i, '') if profil.size > 3
      profil.sub!(/\.dbf$/i, '')
    end
    trc_info :setprof_profilextrahiert, profil

    b_oder_s = bestand_oder_soll.to_s.upcase
    DbfDat::VO_KLASSEN.each { |vklasse|
      prefix = vklasse.name.split("::").last.downcase
      vklasse.set_table_name "" + prefix + b_oder_s + profil #+ ".dbf'"
    }
    #Vp.set_table_name "vp" + $profil
    #Vk.set_table_name "vk" + $profil
  end

  ######################

  class VObjektBasis < ActiveRecord::Base
    def self.vo_sym
      self.name.split("::").last.downcase.to_sym
    end
    def vo_klassen_sym
      self.class.vo_sym
    end
  end

  class VObjektKaskadiert < VObjektBasis
  end

  VObjektBasis.logger = Logger.new(TRACE_DATEI)
  VObjektBasis.colorize_logging = false

  class Vd < VObjektKaskadiert
    set_primary_key :vsnr
    has_one :st, :foreign_key => :vsnr
    has_one :vp, :foreign_key => :vsnr
    has_many :get_vks, :class_name => "Vk" ,:foreign_key => :vsnr

    def vk_haupt
      vks.first
    end

    def vd
      self
    end

private
    def unsortierte_vks
      (@unsortierte_vks = get_vks).each {|vk| vk.vd = self} unless @unsortierte_vks
      @unsortierte_vks
    end

  end


  class St < VObjektKaskadiert  # ZenTest SKIP
    set_primary_key :vsnr
    belongs_to :vd#, :foreign_key => :vsnr
    def st
      self
    end

  end

  class Vp < VObjektKaskadiert  # ZenTest SKIP
    set_primary_key "vsnr"
    belongs_to :vd#, :foreign_key => :vsnr
    def vp
      self
    end

  end

  class Vk < VObjektKaskadiert  # ZenTest SKIP
    belongs_to :get_vd, :class_name => "Vd", :foreign_key => :vsnr
    has_many :get_vvs, :class_name => "Vv" , :foreign_key => [:vsnr, :komp]
    has_many :get_vts, :class_name => "Vt" , :foreign_key => [:vsnr, :komp] #, :conditions => "vtBAT7V4.komp = vkBAT7V4.komp"
    set_primary_keys :vsnr, :komp

    def vts
      if EIGENES_AR_CACHING then
        (@vts = get_vts).each {|vt| vt.vk = self} unless @vts
        @vts
      else
        get_vts
      end
    end

    def vvs
      if EIGENES_AR_CACHING then
        trc_temp "vk-vvs anfang", self.komp
        (@vvs = get_vvs.to_a).each {|vv| vv.vk = self} unless @vvs
        trc_temp "vk-vvs fertig", self.komp
        @vvs
      else
        get_vvs
      end
    end
    
    def vv_max
      @max_vj ||= begin
        my_vvs = vvs.to_a # falls kein EIGENES_AR_CACHING
        max_vj = my_vvs.map{|vv| vv.vj}.max
        my_vvs.find {|vv| max_vj == vv.vj }
      end
    end

    def vv_fuer(vj_oder_zgper)
      case vj_oder_zgper
      when Zeit then
        vvs.to_a.find {|vv| vj_oder_zgper == vv.beg_zeit }
      when Integer
        vvs.to_a.find {|vv| vj_oder_zgper == vv.vj }
      else
        trc_aktuellen_error "keine Zeitangabe: #{vj_oder_zgper.inspect}"
      end
    end

    public
    def vd
      if EIGENES_AR_CACHING then
        @vd ||= get_vd
      else
        get_vd
      end
    end
    attr_writer :vd

    def vt1
      vts.first
    end

    def vk
      self
    end

    def vp
      vd.vp
    end

    def st
      vd.st
    end

  end

  class Vt < VObjektKaskadiert  # ZenTest SKIP
    belongs_to :get_vk, :class_name => "Vk", :foreign_key => [:vsnr, :komp]
    set_primary_keys [:vsnr, :komp, :vtnr]

    def vk
      if EIGENES_AR_CACHING then
        @vk ||= get_vk
      else
        get_vk
      end
    end
    attr_writer :vk


    has_many :get_rks, :class_name => "Rk", :foreign_key => [:vsnr, :komp, :vtnr]
    def rks
      if EIGENES_AR_CACHING then
        (@rks = get_rks).each {|rk| rk.vt = self} unless @rks
        @rks
      else
        trc_temp :vor_getrks
        a = get_rks
        trc_temp :nach_getrks
        a
      end
    end

    has_many :get_vbs, :class_name => "Vb", :foreign_key => [:vsnr, :komp, :vtnr]
    def vbs
      if EIGENES_AR_CACHING then
        (@vbs = get_vbs).each {|vb_1| vb_1.vt = self} unless @vbs
        @vbs
      else
        trc_temp :vor_getvbs
        a = get_vbs
        trc_temp :nach_getvbs
        a
      end
    end

    has_many :get_vfs, :class_name => "Vf", :foreign_key => [:vsnr, :komp, :vtnr]
    def vfs
      if EIGENES_AR_CACHING then
        (@vfs = get_vfs).each {|vf_1| vf_1.vt = self} unless @vfs
        @vfs
      else
        trc_temp :vor_getvfs
        a = get_vfs
        trc_temp :nach_getvfs
        a
      end
    end

    def va
      return @va if @va
      @va = begin
        Va.find( [vk.vd.vsnr, vk.gv, tarif, lfdkz])
      rescue ActiveRecord::RecordNotFound
        nil
      end
    end

    def lfdkz
      if tarif =~ /^FB([EL]).$/ then
        $1
      else
        vk.vd.beiart > 0 ? "L" : "E"
      end
    end

#    def vt1
#      if EIGENES_AR_CACHING then
#        @vt1 ||= get_vt1
#      else
#        get_vt1
#      end
#    end
#
    def vt
      self
    end

    def vd
      vk.vd
    end

    def vp
      vd.vp
    end

    def st
      vd.st
    end

    def beg_zeit
      Zeit.jm(begj, begm)
    end
    
    def rel_fkabw
      (vd.fkabw - vt.begm) % 12
    end
    
    def voriger_rpkt(t)
      if t.m > self.rel_fkabw then
        Zeit.jm(t.j, rel_fkabw)
      else
        if t.j > 0 then
          Zeit.jm(t.j - 1, rel_fkabw)
        else
          Zeit.jm(0,0)
        end
      end
    end
        
      

    def rk_nach_zeit(zeit_relativ)
      if !@rk_nach_zeit
        init_verbindung_rk
      end
      erg = @rk_nach_zeit[zeit_relativ]
      if !erg
        trc_hinweis "rk_nach_zeit ist nil bei zeit=", zeit_relativ
        #trc_info "@rk_nach_zeit", @rk_nach_zeit
      end
      erg
    end

    def init_verbindung_rk
      @rk_nach_zeit = {}
      trc_info "vt-rksnachzeit rks-length:",rks.length
      rks.each_with_index {|rk, idx|
        rk.vt_index = idx
        @rk_nach_zeit[z=Zeit.jm(rk.vj, rk.vm)] = rk;
        trc_temp "z#{z}"
      }
      init_verbindung_vb
    end

    def init_verbindung_vb
      vbs.each do |vb|
        begin
          rk = @rk_nach_zeit[Zeit.jm(vb.vj,vb.vm)]
          rk.vb = vb
          vb.rk = rk
        rescue
          trc_aktuellen_error :verb_vb, 6
        end
      end
    end

  end

  class Rk < VObjektKaskadiert  # ZenTest SKIP
    belongs_to :get_vt, :class_name => "Vt", :foreign_key => [:vsnr, :komp, :vtnr]
    set_primary_keys [:vsnr, :komp, :vtnr, "vj", "vm"]

    def vt
      if EIGENES_AR_CACHING then
        @vt ||= get_vt
      else
        get_vt
      end
    end

    attr_writer :vt_index
    def vt_index
      vt.init_verbindung_rk if not @vt_index
      @vt_index
    end

    attr_accessor :vb

    def vk
      vt.vk
    end

    def vd
      vt.vk.vd
    end

    def rk
      self
    end

    attr_writer :vt
  end

  class Vb < VObjektKaskadiert  # ZenTest SKIP
    belongs_to :get_vt, :class_name => "Vt", :foreign_key => [:vsnr, :komp, :vtnr]
    set_primary_keys [:vsnr, :komp, :vtnr, "vj", "vm"]
#    attr_accessor :vt_index

    def vt
      if EIGENES_AR_CACHING then
        @vt ||= get_vt
      else
        get_vt
      end
    end

    attr_accessor :rk

    def vk
      vt.vk
    end

    def vd
      vt.vk.vd
    end

    def vb
      self
    end

    attr_writer :vt
  end


  class Vv < VObjektKaskadiert  # ZenTest SKIP
    belongs_to :vk, :foreign_key => [ :vsnr, :komp]
  #  has_many :rks
    set_primary_keys [ :vsnr, :komp, :vj]

    def beg_zeit
      Zeit.jm(begj, begm)
    end
    
  end

  class Va < VObjektBasis  # ZenTest SKIP
            #VObjektKaskadiert #VObjektBasis
    has_many :get_vts, :class_name => "Vt", :foreign_key => [:vsnr, :tarif ]
    set_primary_keys [:vsnr, :gv, :tarif, :lfdkz]

  end

  class Vf < VObjektKaskadiert # ZenTest SKIP
    belongs_to :get_vt, :class_name => "Vt", :foreign_key => [:vsnr, :komp, :vtnr]
    set_primary_keys [:vsnr, :komp, :vtnr, "vj", "vm"]
#    attr_accessor :vt_index

    def vt
      if EIGENES_AR_CACHING then
        @vt ||= get_vt
      else
        get_vt
      end
    end

    attr_accessor :rk

    def vk
      vt.vk
    end

    def vd
      vt.vk.vd
    end

    def vf
      self
    end

    attr_writer :vt
    
    def beg_zeit
      Zeit.jm(begj, begm)
    end

  end

  class Gz < VObjektBasis  # ZenTest SKIP
            #VObjektKaskadiert #VObjektBasis
  end

  class Zs < VObjektBasis  # ZenTest SKIP
            #VObjektKaskadiert #VObjektBasis
  end


  VO_KLASSEN = [Vd, St, Vp, Vk, Vt, Rk, Vv, Va, Vb, Vf, Gz, Zs]

end

end # if not defined? DbfDat 

if __FILE__ == $0 then
  durchlaufe_unittests($0)
end
