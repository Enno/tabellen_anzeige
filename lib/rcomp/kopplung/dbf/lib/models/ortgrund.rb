
if not defined? Ort_Grund then

require 'schmiedebasis'

def normalisiere_pfad(pfad)
  pfad.gsub(/[\/\\]+/, "/").chomp("/")+"/"
end


def ats_sort_by(obj)
  name = if obj.respond_to?(:name) then
    obj.name
  else
    obj
  end.to_s.downcase

  if name[0,1] == "<" then
    case name
    when /<excel/  then    "A"
    when /<gftest/ then    "B"
    else                   "D"
    end
  else
                           "C"
  end + name
end

class Ort_Grund
  attr_reader   :vater
  attr_reader   :name
  attr_reader   :anzeige_pfad
  attr_accessor :iterator_eintrag

 # @@ortpool = Hash.new
  #@@speicher_pfad = KONFIG_DIRNAME


#  def self.pfad_zu_ort(pfad)
#    npfad = normalisiere_pfad(pfad)
#    if @
#  end
#
#  def self.normalisiere_pfad(pfad)
#    pfad.gsub(/[\/\\]+/, "/").chomp("/")+"/"
#  end

  def initialize(vater, name)
    @vater   = vater
    @unterorte_hash = {}
    @name = name
    #@ordner_pfad = ordner_pfad
    if vater then
      @anzeige_pfad = vater.anzeige_pfad + name
      vater.neuer_unterort(self)
    else
      @anzeige_pfad = name
    end
    @anzeige_pfad += "/" unless @anzeige_pfad =~ /\/$/
  end

  def komplett_aufloesen
    loesche_unterorte
    if @vater then
      @vater.entferne(self)
    end
    @vater = nil
  end

  def loesche_unterorte
    unterorte.each do |unterort|
      unterort.komplett_aufloesen
    end
    @unterorte_hash = {}
  end

  def neuer_unterort(ort)
    @unterorte_hash[ort.name] = ort
  end

  def entferne(ort)
    trc_temp :entf_ort, ort.name
    @unterorte_hash.delete(ort.name)
  end

  def dienstfokus=(neu_zustand)
    @gerade_im_dienst = neu_zustand
  end

  def xxx_extrahiere_name(dateiname)
    dateiname
  end

  def system_pfad  
    self.anzeige_pfad
  end
  
  def ordner_pfad
    self.system_pfad
  end
  
  alias :kanonischer_pfad :anzeige_pfad

  def to_s
    anzeige_pfad
  end


  def unterorte_erlaubt?
    false
  end

  def unterort_per_name(name)
    @unterorte_hash[name]
  end

  def unterorte
    @unterorte_hash.values.sort_by {|o| ats_sort_by(o.name)}
  end


  alias :subplaces :unterorte

  alias :parent_ort :vater


  def kurzinfo(eintrag) ; nil ; end

  def setze_kurzinfo(eintrag, wert) ; nil ; end

  def quell_ort  ; nil ; end

  def quell_ort=(wert)  ; nil ; end

  def ziel_ort  ; nil ; end

  def ziel_ort=(wert)  ; nil ; end



  def inspect
    "#<#{self.class.name}:#{self.object_id} pfad=\"#{self.anzeige_pfad}\">"
  end

  def direkter_unterort_per_pfad(gesuchter_pfad)
    eigener_pfad = anzeige_pfad
    speztrc = ((eigener_pfad+gesuchter_pfad) =~ /xls/)
    if speztrc then
      trc_temp "%%%%%% gesucht:", gesuchter_pfad
      trc_temp "%%%%%% eigen:  ", eigener_pfad
    end
    return nil if eigener_pfad != gesuchter_pfad[0, eigener_pfad.size]

    rest_gesuchter_pfad = gesuchter_pfad[eigener_pfad.size .. -1]
    if rest_gesuchter_pfad =~ /\// then
      next_stueck = rest_gesuchter_pfad.split("/",2)[0]
    else
      next_stueck = rest_gesuchter_pfad # ### extrahiere_name(rest_gesuchter_pfad)
    end
    if speztrc then
      trc_temp "%% rest:", rest_gesuchter_pfad
      trc_temp "%% next:", next_stueck
    end
    unterort_per_name(next_stueck)
  end

  def finde_ort_per_pfad(gesuchter_pfad)
    trc_temp "suche:   ",  gesuchter_pfad
    trc_info "such-anf:",  anzeige_pfad
    trc_temp "start_ort", self
#    return nil if empty?
    #gesuchter_pfad += '/' if gesuchter_pfad[-1,1] != '/'

    if anzeige_pfad.chomp("/") == gesuchter_pfad.chomp("/")
      trc_info :gefunden, anzeige_pfad
      return self
    end

    next_ort = direkter_unterort_per_pfad(gesuchter_pfad)
    if next_ort then
      return next_ort.finde_ort_per_pfad(gesuchter_pfad)
    end
    nil # vielleicht noch eine Suche einbauen für den Fall dass next_stueck selbst einen slash enthält??
  end

end

end # if not defined? Ort_Grund 

if __FILE__ == $0
  durchlaufe_unittests($0)
end
