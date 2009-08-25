class InfoController < ApplicationController
  set_model 'InfoModel'
  set_view 'InfoView'
  set_close_action :exit

  add_listener :type => :mouse, :components => ["button1_button"]

  def button1_button_action_performed
    close
  end

  def set_label(dialog_text)
    dialog_text.each do |dialog_element, text|
      model.send("#{dialog_element}=", text)
    end
    update_view
  end
end
