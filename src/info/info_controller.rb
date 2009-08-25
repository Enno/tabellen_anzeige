class InfoController < ApplicationController
  set_model 'InfoModel'
  set_view 'InfoView'
  set_close_action :exit


  add_listener :type => :mouse, :components => ["ok_button"]

  def ok_button_action_performed
    close
  end

  def set_label(label)
    model.message = label
    update_view
  end

end
