class InfoView < ApplicationView
  set_java_class 'info.InfoDialog'

  def create_main_view_component
    InfoDialog.new(nil, true)
  end

  def load

  end

  map :view => "text_label.text",     :model => :label
  map :view => "button1_button.text", :model => :button1_text
end
