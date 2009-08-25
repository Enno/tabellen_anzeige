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

  def set_label(label, button1_text, button2_text)
    model.label = label
    model.button1_text = button1_text
    model.button2_text = button2_text
    update_view
  end

  def dialog_result
    model.dialog_result
  end
end
