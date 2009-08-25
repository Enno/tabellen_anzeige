class BestaetigungController < ApplicationController
  set_model 'BestaetigungModel'
  set_view 'BestaetigungView'
  set_close_action :hide

  add_listener :type => :mouse, :components => ["button1_button", "button2_button"]

  def button1_button_action_performed
    model.dialog_result = true
    update_view
    close
  end

  def button2_button_action_performed
    model.dialog_result = false
    update_view
    close
  end

  def set_label(dialog_text)
    dialog_text.each do |dialog_element, text|
      model.send("#{dialog_element}=", text)
    end
    update_view
  end

  def dialog_result
    model.dialog_result
  end
end
