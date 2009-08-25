#require "set"

require 'win32ole'
#require 'excel-library'
require 'schmiedebasis'
require 'models/system_fkt/speicherung'

PROGRAMM_MIT_DBF_ZUGRIFF = false unless defined? PROGRAMM_MIT_DBF_ZUGRIFF

if PROGRAMM_MIT_DBF_ZUGRIFF #and not IST_DIENER_PROZESS then
  require 'models/system_fkt/dbase_zugriff'
end

if not defined? ExcelZugriff then
  
XLCellTypeLastCell = 11
XLWhole            = 1
XLPart             = 2
XLDown             = -4121 
XLShiftDown        = XLDown
XLShiftToRight     = -4161
XLPasteValues      = -4163
XLValues           = -4163 
XLFormulas         = -4123


module ExcelApplicationHelfer
  def Visible=(neu_visible)
    neu_visible = (not (not (neu_visible)))
    #trace "[neu_visible, @@xlApp.Visible]", [neu_visible, @@xlApp.Visible]
    if neu_visible ^ self.Visible then
      super neu_visible
      if neu_visible
        # workaround für unsichtabrkeitsproblem
        self.DisplayFullScreen = true
        self.DisplayFullScreen = false
      end
      trc_info "Excel.visible geändert:",  self.Visible
    end
  end
end  
  
class ExcelZugriff
  @@sichtbar = nil
  @@verbundene_instanzen = [] # zuletzt erzeugte steht zuerst

  attr_reader :app

  def self.aktive_oder_neue_instanz
    ez = lebende_verbundene_instanz
    if ez then
      ez
    else
      verbinde_oder_neu
    end
  end
  
  def self.laufende_instanz_oder_nil
    lebende_verbundene_instanz || begin
      verbinde_mit_laufendem_excel
    rescue
      trc_aktuellen_error "verbinde mit laufendem", 6
      nil
    end
  end

  def self.lebende_verbundene_instanz
    loop do
      break nil if @@verbundene_instanzen.empty?
      ez = @@verbundene_instanzen.first
      if ez.lebt? then
        break ez
      else
        @@verbundene_instanzen.shift
      end
    end
  end

  def self.verbinde_oder_neu(optionen={})
    ez = begin
      verbinde_mit_laufendem_excel(optionen)
    rescue
      neu_erzeugtes_excel(optionen)
    end
    #realisiere_sichtbarkeit
    ez
  end

  def self.finde_instanz(hwnd) # #+#
    @@verbundene_instanzen.find do |ez|
      hwnd == ez.app.HWnd  rescue  nil
    end
  end

  def self.verbundene_instanzen
    @@verbundene_instanzen
  end
  
  def self.verbinde_mit_laufendem_excel(optionen={})
    app = WIN32OLE.connect('Excel.Application')
    trc_hinweis "Verbinde mit schon laufendem Excel, App=", app
    
    # In diesem Fall kann es sein, dass dieses Excel aus irgendwelchen Grï¿½nden hï¿½ngt
    begin
      visb = app.Visible # irgendwas, um zu prüfen, ob Excel ansprechbar ist
    rescue WIN32OLERuntimeError
      raise ExcelErrorBlockiert, "Kann nicht auf Excel zugreifen."
    end
    
    ez = finde_instanz(app.HWnd)
    if ez then
      trc_info :ez_gefunden, ez
    else
      trc_info :ez_nicht_gefunden
      ez = new(app, optionen) 
    end
    ez
  end

  def self.neu_erzeugtes_excel(optionen={})
    app = WIN32OLE.new('Excel.Application')
    trc_hinweis "Neues Excel erzeugt. App=", app
    new(app, optionen)
  end

  def initialize(w32ole_excel_app, optionen={})
    optionen = {:sichtbar=>@@sichtbar}.update(optionen)
    @app = w32ole_excel_app
    @app.extend ExcelApplicationHelfer
    @hwnd = @app.HWnd
    self.class.registriere(self)
    self.visible = optionen[:sichtbar] unless optionen[:sichtbar].nil?
    trc_info :app_hinstance, (["%x"%@app.HInstance] rescue "!Fehler!")
    trc_info :app_HWnd, (@app.HWnd rescue "!Fehler!")
  end

  def self.registriere(ez)
    @@verbundene_instanzen.unshift(ez)
    @@xlApp = ez.app
  end

  def lebt?
    if @app then
      begin
        visb = @app.Visible # irgendwas, um zu sehen ob ein Fehler ausgelöst wird. Nebeneffekt mit Sichtbarkeitstest ist ï¿½berflï¿½ssig
        trc_hinweis "Excelzugriffs-Objekt #{@app} noch intakt, Visible=", visb
        true
      rescue
        trc_hinweis "Excelzugriffs-Objekt #{@app} funktioniert nicht mehr"
        nil
      end
    end
  end

  def visible=(neu_visible)
    neu_visible = (not (not (neu_visible)))
    #trace "[neu_visible, @@xlApp.Visible]", [neu_visible, @@xlApp.Visible]
    if neu_visible ^ @app.Visible then
      @app.Visible = neu_visible
      if neu_visible
        # workaround für Unsichtabrkeitsproblem
        @app.DisplayFullScreen = true
        @app.DisplayFullScreen = false
      end
      trc_info "Excel.visible geändert:",  @app.Visible
    end
  end

  def visible
    @app.Visible
  end


  def self.application
    ez = aktive_oder_neue_instanz

    @@xlApp
  end

  def ExcelZugriff.application_old
    if @@xlApp then
      begin
        visb = @@xlApp.Visible # irgendwas, um zu sehen ob ein Fehler ausgelï¿½st wird. Nebeneffekt mit Sichtbarkeitstest ist ï¿½berflï¿½ssig
        trc_hinweis "Excelzugriffs-Objekt noch intakt, Visible=", visb
      rescue
        trc_hinweis "Excelzugriffs-Objekt funktioniert nicht mehr"
        @@xlApp = nil
      end
    end
    if not @@xlApp then
      begin
        @@xlApp = WIN32OLE.connect('Excel.Application')
        existierendes_excel_benutzt = true
        trc_hinweis "Verbinde mit schon laufendem Excel App=", @@xlApp
      rescue
        @@xlApp = WIN32OLE.new('Excel.Application')
        existierendes_excel_benutzt = false
        trc_hinweis "Neues Excel erzeugt. App=", @@xlApp
      end

      if existierendes_excel_benutzt
        # In diesem Fall kann es sein, dass dieses Excel aus irgendwelchen Grï¿½nden hï¿½ngt
        begin
          visb = @@xlApp.Visible # irgendwas, um zu prï¿½fen, ob Excel ansprechbar ist
        rescue WIN32OLERuntimeError
          raise ExcelErrorBlockiert, "Kann nicht auf Excel zugreifen."
        end
      end
      realisiere_sichtbarkeit

      def @@xlApp.Visible=(neu_visible)
        neu_visible = (not (not (neu_visible)))
        #trace "[neu_visible, @@xlApp.Visible]", [neu_visible, @@xlApp.Visible]
        if neu_visible ^ @@xlApp.Visible then
          super neu_visible
          if neu_visible
            # workaround für unsichtabrkeitsproblem
            @@xlApp.DisplayFullScreen = true
            @@xlApp.DisplayFullScreen = false
          end
          trc_info "Excel.visible geändert:",  @@xlApp.Visible
        end
      end

      trc_info :app_hinstance, (["%x"%@@xlApp.HWnd] rescue "!Fehler!")

    end
    @@xlApp
  end
  
  def self.ohne_displayalerts(mappe_oder_app=nil)
    if mappe_oder_app then
      app = (mappe_oder_app.Application rescue mappe_oder_app)
    else
      app = @@xlApp
    end
    app.DisplayAlerts = false
    begin
      yield
    ensure
      app.DisplayAlerts = true
    end   
  end
  
  

  def self.sichtbar
    @@sichtbar
  end

  def self.sichtbar=(neue_sichtbarkeit)
    @@sichtbar = neue_sichtbarkeit
    realisiere_sichtbarkeit
  end

  def self.sichtbarkeit_reparieren
    ez = laufende_instanz_oder_nil
    if ez then
      a = ez.app
      a.DisplayFullScreen = true
      a.DisplayFullScreen = false
    end
  end


  def self.realisiere_sichtbarkeit
    trc_info :Excelzugriff_sichtbar, @@sichtbar
    exl = laufende_instanz_oder_nil
    if nil != @@sichtbar and exl then
      exl.app.Visible = @@sichtbar
      sichtbarkeit_reparieren
    end
  end

  def self.als_vorderstes_fenster
    application.Visible = true
    SetForegroundWindow.call(application.HWnd)
  end

  def self.schlieszen(aktion_falls_ungespeicherte_mappe=:raise)
    ez = @@verbundene_instanzen.first
    ez.schlieszen(aktion_falls_ungespeicherte_mappe) if ez
  end


  def self.ungespeicherte_mappen_vorhanden?
    ez = lebende_verbundene_instanz
    if not ez then # der Ablauf ist wie bei self.aktive_oder_neue_instanz, auï¿½er dass keine neue Excel-Instanz erzeugt wird
      ez = begin
        verbinde_mit_laufendem_excel
      rescue
        return false
      end
    end
    not ez.ungespeicherte_mappen.empty?
  end

  def ungespeicherte_mappen
    umappen = []
    @app.Workbooks.each do |mappe|
      umappen << mappe if not mappe.Saved
    end
    umappen
  end

  # aktion_falls_ungespeicherte_mappe kann sein :raise, :excel, :forget oder :accept
  def schlieszen(aktion_falls_ungespeicherte_mappe=:raise)

    return if not self.lebt?
    trc_info :app_hwnd, @app.HWnd rescue nil
    weak_wkbks = @app.Workbooks
    trc_temp :weak_wkbks, weak_wkbks

    if not ungespeicherte_mappen.empty? then
      case aktion_falls_ungespeicherte_mappe
      when Proc then
        aktion_falls_ungespeicherte_mappe.call(self, ungespeicherte_mappen)
      when :raise then
        raise "Excel wird nicht beendet, da noch ungespeichtere Mappen offen sind."
      when :excel then
        #nix, Excel wird Fragen stellen
      when :forget then
        ungespeicherte_mappen.each {|m| m.Saved = true}
      when :accept then
        ungespeicherte_mappen.each {|m| m.Save}
      else
        raise "Unbekannte option für 'aktion_falls_ungespeicherte_mappe' (#{aktion_falls_ungespeicherte_mappe})"
      end
    end

    @app.Workbooks.Close rescue trc_hinweis :WorkbooksClose, $!
    weak_wkbks = nil
    weak_wkbks = @app.Workbooks
    trc_temp :weak_wkbks, weak_wkbks
    weak_wkbks = nil
   end


  def self.alle_beenden(aktion_falls_ungespeicherte_mappe=:raise, &blk)
    aktion_falls_ungespeicherte_mappe = blk if blk

    @@verbundene_instanzen.each do |ez|
      #ez.app.Interactive = false
      ez.app.DisplayAlerts = false
    end

    beendet_anzahl = error_anzahl = gesamt_anzahl = 0
    erster_error = nil
    beende_aktion = proc do |ez|
      begin
        gesamt_anzahl += 1
        trc_speicher "#{gesamt_anzahl}}"
        beendet_anzahl += ez.beenden(aktion_falls_ungespeicherte_mappe)
      rescue
        erster_error = $!
        trc_aktuellen_error :fehler_beim_beenden, 7
        error_anzahl += 1
      end
    end

    20.times do
      break if @@verbundene_instanzen.empty?
      beende_aktion.call(@@verbundene_instanzen.first)
      break if error_anzahl > 4
    end
    
    alte_error_anzahl = error_anzahl
    9.times do |lauf_nr|
      ez = begin
        new(WIN32OLE.connect('Excel.Application'))
      rescue
        nil
      end
      beende_aktion.call(ez) if ez
      ole_objekte_zeigen_und_freigeben(true)
      break if not ez
      break if error_anzahl > alte_error_anzahl + 3 
    end
    
    if aktion_falls_ungespeicherte_mappe == :raise and erster_error then
      raise erster_error
    end
    
    [beendet_anzahl, error_anzahl]
  end

  # Beeandet die aktuelle Excel-Applikation
  # aktion_falls_ungespeicherte_mappe:: :raise, :excel, :forget oder :accept
  def self.beenden(aktion_falls_ungespeicherte_mappe=:raise)
    ez = lebende_verbundene_instanz
    ez.beenden(aktion_falls_ungespeicherte_mappe) if ez
  end

  # aktion_falls_ungespeicherte_mappe kann sein :raise, :excel, :forget oder :accept
  def beenden(aktion_falls_ungespeicherte_mappe=:raise)
    require 'weakref'
     #zeige_ole_objekte
     #GC.start
     #zeige_ole_objekte
    beenden_lebender_instanz = self.lebt?
    if beenden_lebender_instanz then
      trc_info :exlquit_app, @app
      hwnd = (@app.HWnd rescue nil)
      trc_info :app_hinstance, hwnd
      schlieszen(aktion_falls_ungespeicherte_mappe)
      @app.Quit
       #trc_temp :weak_wkbks_alive, (defined?(weak_wkbks) and weak_wkbks.weakref_alive?)
      if false and defined?(weak_wkbks)  and weak_wkbks.weakref_alive? then
        trc_temp :weak_wkbks, weak_wkbks
        weak_wkbks.ole_free
      end
       #weak_wkbks = nil
      weak_xlapp = WeakRef.new(@app)
    else
      weak_xlapp = nil
    end # lebt?

    trc_info :delete_erg, @@verbundene_instanzen.delete(self)
    @app = nil
    trc_speicher :_vor_gc
    GC.start
    trc_speicher :nach_gc


    if beenden_lebender_instanz then
      if hwnd then
        pid = get_window_pid(hwnd)
        trc_hinweis :kille_pid, pid
        begin
          trc_hinweis :kill_ok, Process.kill("KILL", pid)
        rescue
          trc_hinweis :kill_error, $!
        end
      end
      if weak_xlapp.weakref_alive? then
         #if WIN32OLE.ole_reference_count(weak_xlapp) > 0
        begin
          trc_temp :weakref_olefree, weak_xlapp.ole_free
        rescue
          trc_aktuellen_error :weakref_probl_olefree
        end
      end
    end

    trc_hinweis "Excel beendet."
    weak_xlapp ? 1 : 0
  end


  def self.ole_objekte_zeigen_und_freigeben(freigeben=true)
    trc_info "-----------ole_objekte---"
    anz_objekte = 0
    ObjectSpace.each_object(WIN32OLE) do |o|
      anz_objekte += 1
      trc_temp :ole_object_name, [o, (o.Name rescue nil)]
      #trc_info :ole_type, o.ole_obj_help rescue nil
      #trc_info :obj_hwnd, o.HWnd rescue   nil
      #trc_info :obj_Parent, o.Parent rescue nil

      if freigeben
        begin
          trc_temp :freed, o.ole_free
        rescue
          trc_hinweis :olefree_error, $!
        end
      end

    end
    trc_hinweis :anz_w32obj, anz_objekte
  end

#######################################################
  
  def sicheres_lesen(bezug, fehlermeldung = "Bezug #{bezug} wurde nicht gefunden")
    begin
      app.Range(bezug).Value
    rescue WIN32OLERuntimeError
      trc_aktuellen_error "lesen von #{bezug}"
      raise fehlermeldung    
    end    
  end
  
  def sicheres_schreiben(bezug, wert)
    app.Range(bezug).Value = wert
  rescue WIN32OLERuntimeError
    trc_aktuellen_error "schreiben: '#{bezug}'=#{wert.inspect}", 4    
  end
  
  
  def self.sicheres_lesen(bezug, fehlermeldung = "Bezug #{bezug} wurde nicht gefunden")
    ez = ExcelZugriff.laufende_instanz_oder_nil
    if not ez then
      raise "Kein zugreifbares Excel geöffnet, also: " + fehlermeldung
    end
    ez.sicheres_lesen(bezug, fehlermeldung)
  end

  def self.sicheres_schreiben(bezug, wert)
    ez = ExcelZugriff.laufende_instanz_oder_nil
    if not ez then
      trc_fehler "Kein zugreifbares Excel geöffnet, also kann nicht in '#{bezug}' den wert #{wert.inspect} schreiben"
      raise "Bug: Kein zugreifbares Excel geöffnet, also kann nicht in '#{bezug}' den wert #{wert.inspect} schreiben"
    else
      ez.sicheres_schreiben(bezug, wert)
    end  
  end


end








#################

class ExcelError < RuntimeError
end

class ExcelErrorBlockiert < ExcelError
end




def erster_positiver_wert(*werte)
  werte.find { |wert| wert > 0 } || 0
end




class ExcelErrorFalscherParameter < ExcelError
end


end # if not defined? ...




if __FILE__ == $0 then

  durchlaufe_unittests($0)

end


