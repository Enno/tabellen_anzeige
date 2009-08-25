#!ruby
# Sven Suska
# encoding: iso-8859-15


require File.dirname(__FILE__) + "/testhilfe"

if __FILE__.gsub("\\", "/") == $0.gsub("\\", "/") and not defined? TESTSTARTED_SCHMIEDEBASIS_RB
  p __FILE__ 
  durchlaufe_unittests($0) 
  
elsif defined?(SCHMIEDEBASIS_GELADEN) then
  # Absicherung gegen doppeltes Laden
  trc_info "****************************************"
  trc_fehler :schmiedebasis_doppelt_GELADEN, SCHMIEDEBASIS_GELADEN
  PROTOKOLLIERER_HAUPT.puts caller
else
  SCHMIEDEBASIS_GELADEN = true

  $: << File.expand_path(File.dirname(__FILE__))
  
  require 'version'
  SCHMIEDEBASIS_VERSION = "1.2.0"
  RELEASE_FFM = false unless defined?(RELEASE_FFM)

  $orig_stdout, $orig_stderr = $stdout,  $stderr

  IST_DIENER_PROZESS = false unless defined?(IST_DIENER_PROZESS)


  if not defined?(TAR2RUBYSCRIPT)
    require 'oldandnewlocation.rb'
  end

  if not defined?(WORK_DIRNAME) then
    WORK_DIRNAME = if $0 =~ /TestRunner\.rb$/ then # Hack für Eclipse-TestUnit-Unterstützung
      File.dirname(__FILE__).gsub("\\","/")
    else
      oldlocation {
        # #*# Ich trau mich nicht, Dir.getwd einzusetzten, aber es sähe schöner aus. Irgenwann mal probieren
        #`cd`.chomp.gsub("\\","/") #workaround funktioniert auch nicht, Dir.getwd bringt Probleme mit langen Dateinamen
        Dir.getwd
      }
    end
  end


  KONFIG_DIRNAME = WORK_DIRNAME + "/.TestSchmiede"
  if not File.exist?(KONFIG_DIRNAME) then
    Dir.mkdir(KONFIG_DIRNAME)
    ##+# versteckten Ordner
  end

  LOG_DIRNAME = WORK_DIRNAME + "/.TestSchmiede"

  require 'models/system_fkt/protokollierer'
  
  tracefile_anlegen = (IST_DIENER_PROZESS || defined?(TAR2RUBYSCRIPT))

  if tracefile_anlegen #or true
    #$trc_dateiname = tracedateiname(IST_DIENER_PROZESS ? "sub" : "main")
    PROTOKOLLIERER_HAUPT = Protokollierer.new("", IST_DIENER_PROZESS)
    PROTOKOLLIERER_HAUPT.setze_limits( 
      :einzel_fixiert_anzahl => 5, 
      :einzel_freigabe_ab_anzahl => 30, 
      :einzel_freigabe_ab_bytes => 2_000_000,      
      :paket_freigabe_ab_bytes => 2_000_000,      
      :paket_datei_bytes => 500_000     
    )
  else # wenn in der Entwicklung befindlich, dann Protokollierer auf Konsole:
    PROTOKOLLIERER_HAUPT = Protokollierer.new(:stderr, IST_DIENER_PROZESS)
  end
  TRACE_DATEI = PROTOKOLLIERER_HAUPT.datei
    
#  if not IST_DIENER_PROZESS
    PROTOKOLLIERER_PERM = Protokollierer.new("per", IST_DIENER_PROZESS)
    
    PROTOKOLLIERER_PERM.setze_limits( 
      :einzel_fixiert_anzahl => 3, 
      :einzel_freigabe_ab_anzahl => 80, 
      :einzel_freigabe_ab_bytes => 500_000,      
      :paket_freigabe_ab_bytes => 1_000_000,      
      :paket_datei_bytes => 200_000     
    )
      
 # end     



  
  # Wenn wir in einem Diener-Prozess sind, brauchen wir die Ein- und Ausgabe zur Kommunikation
  # ansonsten leiten wir alle Ausgabe-IO-Ströme auf die Log-Datei, vorausgesetzt, dass wir überhaupt tracen:    
  $stdout = $stderr = TRACE_DATEI if tracefile_anlegen and not IST_DIENER_PROZESS  

  def trc_temp(wo, was=:nix_und_gar_nix, &blk)
    PROTOKOLLIERER_HAUPT.trace_allgemein(:temp, wo, was, &blk)
  end

  def trc_info(wo, was=:nix_und_gar_nix, &blk)
    PROTOKOLLIERER_HAUPT.trace_allgemein(:info, wo, was, &blk)
  end

  def trc_hinweis(wo, was=:nix_und_gar_nix, &blk)
    PROTOKOLLIERER_HAUPT.trace_allgemein(:hinweis, wo, was, &blk)
  end
  def trc_fehler(wo, was=:nix_und_gar_nix, &blk)
    PROTOKOLLIERER_HAUPT.trace_allgemein(:fehler, wo, was, &blk)
  end

  def trc_essenz(wo, was=:nix_und_gar_nix, &blk)
    PROTOKOLLIERER_HAUPT.trace_allgemein(:essenz, wo, was, &blk)
    PROTOKOLLIERER_PERM.trace_allgemein(:essenz, wo, was, &blk)
  end

  def trc_aktuellen_error(text, max_backtrace = -1)
    PROTOKOLLIERER_HAUPT.trace_allgemein(:fehler, text, $!)
    PROTOKOLLIERER_PERM.trace_allgemein(:fehler, text, $!)
    PROTOKOLLIERER_HAUPT.puts $!.backtrace[0..max_backtrace].join("\n")
    PROTOKOLLIERER_PERM.puts $!.backtrace.first
  end

  def trc_internen_fehler(text = "")
    trc_aktuellen_error("BUG:"+text)
  end
  
  def trc_caller(wo="caller", anzahl=6)
    PROTOKOLLIERER_HAUPT.trace_allgemein(:temp, wo) do
      "\n" + caller[4..-1].first(anzahl).join("\n")
    end
  end


  $trace_stufe ||= IST_DIENER_PROZESS ? :info : :temp
#  $trace_stufe ||= IST_DIENER_PROZESS ? :info : :temp

  trc_essenz "Start: ", Time.now.strftime("%Y-%m-%d %H:%M:%S")
  
begin  
  trc_essenz "Prog-Version",
    (defined?(TESTSCHMIEDE_VERSION) ? TESTSCHMIEDE_VERSION : "SCHMIEDEBASIS-"+SCHMIEDEBASIS_VERSION )

  trc_essenz "Start: ", Time.now.strftime("%Y-%m-%d %H:%M:%S")
  if ARGV.size > 0 then # #*# besser irgendein Opt-Parse-Tool nehmen
    ARGV.find do |arg|
      if arg =~ /^--trace=(\w+)/ then
        $trace_stufe = $1.to_sym
      end
    end
  end

  #exit 0

    if RUBY_PLATFORM =~ /wi/i then

  require 'Win32API'

	module WindozePaths
	  GetLongPathName     = Win32API.new('kernel32', 'GetLongPathName', 'PPL', 'L')
	  GetShortPathName    = Win32API.new('kernel32', 'GetShortPathName', 'PPL', 'L')
	
	  PATHNAME_BUFFER = "-"*500
	  def self.long_pathname(dateiname)
	    res = GetLongPathName.call(dateiname, PATHNAME_BUFFER, PATHNAME_BUFFER.size-1)
	    PATHNAME_BUFFER[0,res]
	  end
	
	  def self.short_pathname(dateiname)
	    res = GetShortPathName.call(dateiname, PATHNAME_BUFFER, PATHNAME_BUFFER.size-1)
	    PATHNAME_BUFFER[0,res]
	  end
	end
    else
 	module WindozePaths
	  def self.long_pathname(dateiname)
      dateiname
	  end

	  def self.short_pathname(dateiname)
      dateiname
 	  end
	end
  class Win32API

    def initialize(*args)
      p ["Win32API-init:", args]
    end
    def method_missing(*args)
      p ["Win32API-call", args]
    end
  end
    end

  Dir.chdir(oldlocation)

  #trace :getwd, Dir.getwd
  #trace :old, oldlocation { Dir.getwd }
  #trace :new, newlocation { Dir.getwd }

=begin
  BIN_DIRNAME =
    oldlocation{
      if File.dirname($0) == "."
        p :bin_dir
        `cd`.chomp.gsub("\\","/") #workaround, Dir.getwd bringt Probleme mit langen Dateinamen
      else
        p :bin_file
        WindozePaths::long_pathname(File.dirname($0))
      end
    }
=end
trc_temp "_FILE_",  __FILE__
trc_temp "__$0__",  $0

  
    orig_dir_alt =newlocation {
      if File.dirname(__FILE__) == "."
        trc_hinweis :orig_dir
        `cd`.chomp #workaround, Dir.getwd bringt Probleme mit langen Dateinamen
      elsif File.dirname(__FILE__) =~ /^..\//
        trc_hinweis :orig_parent
        File.expand_path(File.dirname(__FILE__))
      else
        trc_hinweis :orig_file
        WindozePaths.long_pathname(File.dirname(__FILE__))
      end.gsub("\\","/")
  } # .sub(/\/atsSystem\/?$/, "")
  
  ORIG_DIRNAME = WindozePaths.long_pathname(File.expand_path(File.dirname(__FILE__)))
  
  if orig_dir_alt == ORIG_DIRNAME then 
    trc_hinweis "OD OK"
  else
    trc_essenz "!!! OD einfach weicht ab !!!", [orig_dir_alt, ORIG_DIRNAME]
  end
    

  $: << ORIG_DIRNAME
  $:.uniq!
  trc_hinweis :$:, $:

  #p $:
  

  #SHGetFileInfo = Win32API.new("shell32","SHGetFileInfo",["P","I","P","I","I"], "P")
  #require 'windows/file'

  trc_hinweis :ProzessArt, (IST_DIENER_PROZESS ? "Diener" : "HauptProzess")
  trc_hinweis :wd, WORK_DIRNAME
  #trace :bd, BIN_DIRNAME
  trc_hinweis :od, ORIG_DIRNAME
  #trace :ol, WindozePaths::long_pathname(ORIG_DIRNAME )
  trc_temp :file_old, oldlocation {__FILE__}
  trc_temp :file_new, newlocation {__FILE__}
  trc_temp :sp_old, oldlocation {$0}
  trc_temp :sp_new, newlocation {$0}
  trc_temp :getwd, Dir.getwd()
  trc_hinweis :APPEXE,  (RUBYSCRIPT2EXE_APPEXE  if defined? RUBYSCRIPT2EXE_APPEXE)
  trc_hinweis :TEMPDIR, (RUBYSCRIPT2EXE_TEMPDIR if defined? RUBYSCRIPT2EXE_TEMPDIR)
  trc_essenz :ARGV, ARGV.join(" ## ")
  trc_hinweis :RUBY_VERSION, RUBY_VERSION
  trc_hinweis :$:, $:
  trc_hinweis :aufrufer
  PROTOKOLLIERER_HAUPT.puts caller
  ENV.each do|name, wert|
    trc_hinweis name, wert
  end
  RUBYW_INTERPRETER = if defined? RUBYSCRIPT2EXE_TEMPDIR then
    RUBYSCRIPT2EXE_TEMPDIR+"\\bin\\rubyw.exe"
  else
    "rubyw"
  end
  trc_hinweis :RUBYW_INTERPRETER, RUBYW_INTERPRETER
  trc_fehler "rubyw nicht da!!" unless File.exist?(RUBYW_INTERPRETER) #.gsub("\\","/")

  if defined? RUBYSCRIPT2EXE_TEMPDIR then
	  ENV["RUBYLIB"] = RUBYSCRIPT2EXE_TEMPDIR.gsub("\\","/")+"/lib"
	  #ENV.delete "RUBYOPT"
	  ENV["RUBYOPT"] = "-rubygems"
	  trc_temp :______ENV_NACH_SET________
	  ENV.each do|name, wert|
	    trc_temp name+"2", wert
	  end
  end
  
  RAILS_CONNECTION_ADAPTERS = ["apollo_bde"]
  
  #################################

  PeekMessage = Win32API.new("user32", "PeekMessage", ['P'] + ['I']*4, 'I')
  GetMessage  = Win32API.new("user32", "GetMessage",  ['P'] + ['I']*3, 'I')
  SendMessage = Win32API.new("user32", "SendMessage", ['L'] * 4, 'L')
  PostMessage = Win32API.new("user32", "PostMessage", ['L'] * 4, 'I')

  W32MessageBox = Win32API.new('user32', 'MessageBox', 'LPPL', 'I')

  GetWindowThreadProcessId = Win32API.new("user32", "GetWindowThreadProcessId", ["I","P"], "I")
  def get_window_pid(hwnd)
    pid_puffer = " " * 32
    trc_temp :retval, GetWindowThreadProcessId.call(hwnd, pid_puffer)
    pid_puffer.unpack("L")[0]
  end

  GlobalMemoryStatus = Win32API.new('kernel32', 'GlobalMemoryStatus', 'P', 'V')
  def trc_speicher(hinweis)
    PROTOKOLLIERER_HAUPT.trace_allgemein(:info, "FREI-SPEICH #{hinweis}") do
      ms = [32,0,0,0,0,0,0,0].pack("LLLLLLLL")
      GlobalMemoryStatus.call(ms)
      l, mload, tl_phys, av_phys, tl_pf, av_pf, tl_v, av_v = ms.unpack("LLLLLLLL")
      "%2d%% %7d-phys %8d-ausl %8d-max_a"%[mload,av_phys/1024,av_pf/1024,tl_pf/1024]
    end
  end

  trc_speicher :start

  if not IST_DIENER_PROZESS then

    PROGRAMMNAME = "Apollos TestSchmiede" unless defined? PROGRAMMNAME

  if RUBY_PLATFORM =~ /wi/i then

    trc_temp :INIT_PHI_INIT_PHI_INIT_PHI
    begin
      erg = require 'phi'
      trc_hinweis 'phi', (erg ? 'soeben' : 'war bereits') + ' geladen'
      require 'dialogs'
    rescue
      trc_aktuellen_error 'require phi'
      W32MessageBox.call 0, <<TextEnde, "Problem beim Starten", 0
Die Anwendung konnte nicht gestartet werden,
da erforderliche Bibliotheken nicht gefunden wurden.

Apollos Testschmiede setzt eine Installation von
Delphi 6 mit mindestens Update Pack 2 voraus.
TextEnde
    end
    
=begin
    require 'iconv'
    module Phi # ZenTest SKIP
      DELPHI_ENCODING = 'LATIN1' #'UTF-8' # 'ISO-8859-1' #'Shift_JIS' #'ISO-8859-1' #'Shift_JIS'
      RUBY_ENCODING = 'UTF-8' # 'ISO-8859-1' #'UTF-8'
      RubyDelphiCodeConverter = Iconv.new(DELPHI_ENCODING, RUBY_ENCODING)
      DelphiRubyCodeConverter = Iconv.new(RUBY_ENCODING, DELPHI_ENCODING)
    module_function
      def convert_encoding_from_ruby_to_delphi(str)
         begin
          RubyDelphiCodeConverter.iconv(str)
        rescue
          trc_aktuellen_error "!!! CHAR-CONV von unicode FEHLGESCHLAGEN !!!"
          str
        end
      end

      def convert_encoding_from_delphi_to_ruby(str)
        #trc_temp [:delphi_to_ruby, str]
        begin
          DelphiRubyCodeConverter.iconv(str)
        rescue
          trc_aktuellen_error "!!! CHAR-CONV von latin FEHLGESCHLAGEN !!!"
          str
        end
      end
    end
=end

    def yield_to_events
      Phi::Application.instance.process_messages
    end

  end
  else

  end #if not IST_DIENER_PROZESS then

  def excel_zugriff
#    trc_info :exl_require,
      require('models/system_fkt/excel_zugriff')
    trc_temp caller.first(3).join("\n")
    ExcelZugriff
  end


  # hack für activerecord und r2exe
  alias :orig_vor_svens_require :require

if false
  ACTRECOR_PFAD = "C:/ProgLang/RunRuby/lib/ruby/1.8/actrecor"
  def require(pfad, *rest_args, &blk)
    pfad_nachher = pfad.sub(/#{ACTRECOR_PFAD}\//i, "")
    if pfad != pfad_nachher then
    #  trc_temp :pfad_nachher_vorher, [pfad_nachher, pfad]
    end
    orig_vor_svens_require( pfad_nachher, *rest_args, &blk)
  end
end

  class Array
    alias old_case_eql :===
    def ===(wert)
      if wert.is_a? Array then
        old_case_eql(wert)
      else
        include?(wert)
      end
    end
  end
  
  trc_hinweis :SCHMIEDEBASIS_fertig
  
#rescue
#  notfall_reaktion_auf_fehler("schmiedebasis_init")
#  trc_aktuellen_error "schmiedebasis_init"
end
end # unless already loaded
  

