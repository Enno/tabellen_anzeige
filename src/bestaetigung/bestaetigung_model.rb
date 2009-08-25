class BestaetigungModel
  attr_accessor :label, :dialog_result, :button1_text, :button2_text
  def alle_texte=(dialog_texte_hash)
    dialog_texte_hash.each do |was, inhalt|
      send("#{was}=", inhalt)
    end
  end

  def initialize
    @label = "Label"
    @dialog_result = nil
    @button1_text = "Button1"
    @button2_text = "Button2"
  end

end
