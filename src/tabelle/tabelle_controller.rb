require 'spaltenwahl_controller'

class TabelleController < ApplicationController
  set_model 'TabelleModel'
  set_view 'TabelleView'
  set_close_action :exit

  def spaltenwahl_btn_action_performed
    SpaltenwahlController.instance.open
  end
end
