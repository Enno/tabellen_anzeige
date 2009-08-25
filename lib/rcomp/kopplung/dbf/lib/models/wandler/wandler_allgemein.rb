
if not defined? WandlerAllgemein then


require 'schmiedebasis'



class WandlerAllgemein

  attr_reader :waart_str, :waart_obj, :quellort_pfad, :ziel_pfad, :opts

  def initialize(quellort_pfad, ziel_pfad, opts = {})
    @quellort_pfad = quellort_pfad
    @ziel_pfad = ziel_pfad
    @opts = opts
    @waart_str = self.class.name.split("_")[1]
    if not @waart_str then 
      raise AtsBug, "Wandler-Klassenname (#{self.class.name}) entspricht nicht der Konvention für Wandler-Klassen"
    end
    @waart_str.upcase!
    @waart_obj = WANDEL_ARTEN.find {|wa| wa.bez == @waart_str }
  end
  
  def neuer_wandler(wandel_art, quellort_pfad, ziel_pfad, optionen = {})
    waart_str = wandel_art.to_s        
    klassen_name = "Wandler_" + waart_str.downcase.capitalize
    wandel_klasse = const_get(klassen_name)
    wandel_klasse.new(ziel_pfad, optionen)
  end
end


class WandlerVertrZuMappe < WandlerAllgemein
  def initialize(quellort_pfad, ziel_pfad, opts = {}) 
    super
  end
    
end

class WandlerVertrZuZeilen < WandlerAllgemein
  def initialize(quellort_pfad, ziel_pfad, opts = {})    
    super
  end    
end


end # if not defined? WandlerAllgemein