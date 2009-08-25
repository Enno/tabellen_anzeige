#PROZENT_Proc  = proc { |wert| (wert||0) / 100 }
#PROMILLE_Proc = proc { |wert| (wert||0) / 1000 }
#require 'schmiedebasis'
#require 'wandler/dbf2exl'

if not defined? ExcelZuordnungNG1 then

def promille(wert)
  (wert || 0)  /  1000
end

plugin_ordner = File.dirname(__FILE__) + "/excel_zuordnung"
if not File.exist?(plugin_ordner) then
  raise "Bug: Plug-In Ordner existiert nicht! (#{plugin_ordner}"
end

Dir[plugin_ordner + "/zuord_*.rb"].sort.each do |dn|
  require dn
  trc_temp :zuord_required, dn
end

class EZ
  def self.anw(v_object, komp_nr=nil)
    #return
    erg = {}
    prefix = v_object.class.name.split("::").last.upcase
    ExcelZuordnungNG1.const_get(prefix+"_zuord").each do |exlname, exl_felddef|
      #exl_felddef = (self.class)::FELDER_EXL[exlname.to_sym]
      wert =
        case exl_felddef
          when :nix  then next
          when Proc  then v_object.instance_eval(&exl_felddef)
          else            v_object.attributes[exl_felddef.to_s]
        end

      exlname = "#{exlname}_k#{komp_nr}".to_sym if komp_nr
      trc_temp :exlname_wert_def , [exlname, wert, exl_felddef]
      erg[exlname] = wert
    end
    erg
  end

  def eintragen
    eingaben = @mappe.Sheets("Eingaben")
    eingaben.Activate
      begin
        zelle = eingaben.Range(exlname)
      rescue
        next
      end
      excelwert_eintragen(zelle, wert)
    #neuer_bereich(blatt, exlnames)
  end

end
end # if not defined? ExcelZuordnungNG1