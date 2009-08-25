class BestaetigungView < ApplicationView
  set_java_class 'bestaetigung.BestaetigungDialog'

    def create_main_view_component
    BestaetigungDialog.new(nil, true)
  end

  def load
  end

  map :view => "text_label.text", :model => :label
  map :view => "button1_button.text", :model => :button1_text
  map :view => "button2_button.text", :model => :button2_text

end
