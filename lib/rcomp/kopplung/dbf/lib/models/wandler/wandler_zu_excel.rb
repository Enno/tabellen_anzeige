
if not defined? WandlerZuExcel then

require 'schmiedebasis'
PROGRAMM_MIT_DBF_ZUGRIFF = true unless defined? PROGRAMM_MIT_DBF_ZUGRIFF
require 'models/system_fkt/dbase_zugriff'
require 'models/system_fkt/excel_zugriff'
require 'models/generali'
require 'models/wandler/excel_zuordnung'


class WandlerZuExcel
  attr_reader :mappe 
  
  
  def initialize(ziel_pfad, vorlage_mappe)
    ExcelZugriff.speichere_kontrolliert(
      vorlage_mappe,
      ziel_pfad.gsub(/[<>]/,"").chomp("/"),
      KONFIG.opts[:WennBeiExcelErstellungMappeExistiert]
    )
    @mappe = vorlage_mappe 
    @blatt_wandler = []
    mappe.Worksheets.each do |blatt|
      @blatt_wandler << BlattWandler.new(blatt)
    end
  end
  
  def finalisiere
    @blatt_wandler.each do |bw|
      bw.finalisiere
    end
    @mappe.Save if @mappe
  end
  
  def starte_wandlung(vds)
    @blatt_wandler.each do |blw|
      blw.fuelle_blatt(vds)
    end
  end
    
end

class BlattWandler
  attr_reader :akt_blatt

  @@zuord_modul = ExcelZuordnungHR6
    
  def initialize(blatt)
    @akt_blatt = blatt
    @letzte_zeil_nr = 2
#    DbfDat.
  end
  
  def finalisiere
    akt_blatt.Rows(2).RowHeight = 0 if @primkeys and not @primkeys.empty?
  end
       
  def fuelle_blatt(vds)
    #self.akt_blatt = blatt
    #def akt_blatt= blatt
    #@akt_blatt = blatt
    
    ziel_spalten_ermitteln
    unless @primkeys.empty?
      akt_blatt.Activate
      trage_alle_werte_in_blatt_ein(vds)
    end 
  end    
  
  def ziel_spalten_ermitteln
    if @spalten_nr_hash.nil? then
	    @spalten_nr_hash = {}
	
	    spalten_namen = kopfzeile_auswerten
	    
	    DbfDat::VO_KLASSEN.each do |vo_klasse|
	      vo_symbol = vo_klasse.vo_sym
	      zuord = self.class.zuord_defs(vo_symbol)
	      next unless zuord      
	           
	      erg = {}
	      moegl_zielnamen = zuord.keys.map {|n| n.to_s.downcase}
        if vo_klasse.connected? then 
	        moegl_zielnamen = (moegl_zielnamen + vo_klasse.column_names).uniq
        end
	      
		    spalten_namen.each_with_index do |sp_name, idx|
		      sp_nummer = idx + 1
		      #sp_name = sp_ueberschrift.to_s.strip.downcase
		      if moegl_zielnamen.include? sp_name and akt_blatt.Cells(2,sp_nummer).Formula !~ /^=/ then
		        erg[sp_name] = sp_nummer
		      end                     
		    end
	      @spalten_nr_hash[vo_symbol] = erg unless erg.empty?
	    end
    end
    @spalten_nr_hash   
  end
  
  def kopfzeile_auswerten
    b = akt_blatt  
    letzte_spalte = b.UsedRange.Columns.Count
    kopfzeile     = b.Range(b.Cells(1,1), b.Cells(1, letzte_spalte))
    sp_namen      = kopfzeile.Value.first.map {|w| w.to_s.strip.downcase }     
    #muster_zeile  = b.Range(b.Cells(2,1), b.Cells(2, letzte_spalte))
    
    @primkeys = begin
      b.Application.Intersect(b.Range("PrimKeys"), kopfzeile).Value.first
    rescue WIN32OLERuntimeError
      trc_aktuellen_error :kein_primkey, 3
      []
    end.map {|w| w.to_s.strip.downcase }
    
    sp_namen
  end
    
  
  
  attr_accessor :letzte_zeil_nr
  
  def trage_alle_werte_in_blatt_ein(vds)
    #sp_nummern = alle_ziel_spalten_nummern(blatt)
    
    vds.each do |vd|
      vd_eintr = neue_eintraege_fuer(vd)
      st = vd.st
      vd_eintr.update neue_eintraege_fuer(st)
      vd_eintr.update neue_eintraege_fuer(vd.vp)
      vd.vks.each do |vk|
        vk_eintr = vd_eintr.dup.update neue_eintraege_fuer(vk)  
        
=begin  
                  vvs = vk.vts.map do |vt|
                  end.compact.uniq
                  #vvs.each do |vv|
=end
        if @primkeys.include?("vtnr") then
          vk.vts.each do |vt|
            vt_eintr = vk_eintr.dup.update neue_eintraege_fuer(vt)            
            vt_eintr.update neue_eintraege_fuer(vt.va)
            
            if @primkeys.include?("vm") then
              vjm = st.zgper_zeit - vt.beg_zeit # + Zeit.jm(1,0)               
              trc_info :vjm, vjm       
              rk = vt.rk_nach_zeit(vjm)
              vt_eintr.update neue_eintraege_fuer(rk)
#            else              
            end
            fuelle_neue_zeile(vt_eintr)
          end
          
        elsif @primkeys.include?("vj") then          
          vv = if akt_blatt.Name !~ /-ref$/ then
            vt = vk.vt1
            t_zeit = st.zgper_zeit - vt.beg_zeit
            t_zeit = vt.voriger_rpkt(t_zeit)
            vj = (t_zeit - Zeit.new(vt.rel_fkabw)).j + 1
            vk.vv_fuer(vj)
          else
            vk.vv_fuer(st.zgper_zeit)
          end
          
          vv_eintr = vk_eintr.dup.update neue_eintraege_fuer(vv)            
          fuelle_neue_zeile(vv_eintr)
          #end          
          
        
        else          
          fuelle_neue_zeile(vk_eintr)
        end
        
      end
    end
    
  end
  
  def neue_zeile_anhaengen
    @letzte_zeil_nr += 1
    neue_zeile = akt_blatt.Rows(letzte_zeil_nr)
    if @letzte_zeil_nr <= akt_blatt.Rows.Count then
	    neue_zeile.Select
	    akt_blatt.Application.Selection.Insert("Shift" => XLDown)
    end
    
    akt_blatt.Rows(2).Copy
    neue_zeile = akt_blatt.Rows(letzte_zeil_nr)
    neue_zeile.Select
    akt_blatt.Paste
    letzte_zeil_nr
  end
  
  def fuelle_neue_zeile(*eintr_hashs)
    z_nummer = neue_zeile_anhaengen
    #zeile = akt_blatt.Rows(z_nummer)
    eintr_hashs.each do |eintr_hash|
      eintr_hash.each do |sp_nummer, wert|
        self.class.schreibe_zellwert(akt_blatt.Cells(z_nummer, sp_nummer), wert)        
      end
    end    
  end 
  
  def self.schreibe_zellwert(zelle, wert)
    zelle.Value = wert 
    zelle.Interior.ColorIndex = 42
  end
  
  def self.zuord_defs(vo_symbol)
    @@zuord_defs ||= Hash.new 
    unless @@zuord_defs.has_key?(vo_symbol)
      neu_hash = {}        
      orig_hash = (@@zuord_modul.const_get(vo_symbol.to_s.upcase + "_zuord") rescue nil)
      if orig_hash then
        orig_hash.each {|exl_name, definition| neu_hash[exl_name.to_s.downcase] = definition}
      end
      @@zuord_defs[vo_symbol] = neu_hash
    end
    @@zuord_defs[vo_symbol]
  end
  
  def neue_eintraege_fuer(vobject)
    erg = {}
    return erg if vobject.nil?
    
    vo_symbol = vobject.vo_klassen_sym    
    
    spalten_nr_hash = @spalten_nr_hash[vo_symbol]
    return erg if spalten_nr_hash.nil?    
    
    zuordnung = self.class.zuord_defs(vo_symbol)
    
    spalten_nr_hash.each do |exlname, sp_nummer|
      exl_felddef = zuordnung[exlname]
      if not exl_felddef then
        exl_felddef = exlname
        #trc_temp "nicht in exl-zuord gefunden, feld:", exlname
      end
      #exl_felddef = (self.class)::FELDER_EXL[exlname.to_sym]
      wert =
        case exl_felddef
        when :nix  then next
        when Proc  then
          begin
            vobject.instance_eval(&exl_felddef)
          rescue
            trc_aktuellen_error :felddef_eval, 10
            $ats.konsole.meldung "Wert für Zelle:#{exlname} konnte nicht berechnet werden. Fehler=#{$!}"
          end
        else
          begin
            vobject.attributes[exl_felddef.to_s]
          rescue
            trc_aktuellen_error :felddef_ausfeld, 8
            $ats.konsole.meldung "Zelle:#{exlname} ist mit #{vo_symbol}.#{exl_felddef} verknï¿½pft, das Feld wurde aber nicht in der Tabelle gefunden"
          end
        end

      trc_temp :exlname_wert_def , [exlname, wert, exl_felddef]
      
       
      #sp_nummer = spalten_nr_hash[exlname.to_s.downcase]
#      if not sp_nummer then
 #       trc_fehler "Spalten-Nr war nicht gespeichert, für feld:", exlname
  #      next
   #   end
      erg[sp_nummer] = wert
    end
    erg
  end
  
  
end

end # if not defined? WandlerZuExcel 


if __FILE__ == $0 then
  durchlaufe_unittests($0)
end
