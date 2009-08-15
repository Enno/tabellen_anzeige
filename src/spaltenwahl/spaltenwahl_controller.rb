class SpaltenwahlController < ApplicationController
  set_model 'SpaltenwahlModel'
  set_view 'SpaltenwahlView'
  set_close_action :hide

  #add_listener :type => :action, :components => [:ok_btn]

  def spalten_eintragen(hash)
    model.alle_spalten   = hash[:alle]   if hash.has_key?(:alle)
    model.aktive_spalten = hash[:aktive] if hash.has_key?(:aktive)
    update_view
  end

  def ok_btn_action_performed
    p :ok_button_handler_called
    close
  end


  def aktive_spalten
    update_model view_model, :aktive_spalten
    model.aktive_spalten
  end
end
