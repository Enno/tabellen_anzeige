require 'spaltenwahl_controller'
class TabelleController < ApplicationController
  set_model 'TabelleModel'
  set_view 'TabelleView'
  set_close_action :exit

  def spaltenwahl_btn_action_performed
    p @__view.instance_variables
    
    spaltenwahl_controller = SpaltenwahlController.instance
    p :vor_add_nested_contr
    add_nested_controller(:spaltenwahl_dialog, spaltenwahl_controller)
    p :nach_add_nested_contr
#    spaltenwahl_controller.class.add_listener :type => :action, :components => [:ok_btn]
    # hat nicht funktioniert:
    spaltenwahl_controller.define_handler(:ok_btn_action_performed) do
      p :vor_ok_btn_update
    end
    # hat auch nicht funktioniert:
    spaltenwahl_controller.define_handler(:window_closed) do
      p :window_closed
      p :vor_ok_btn_update
    end
    spaltenwahl_controller.open
    spalten = spaltenwahl_controller.ergebnis
    p spalten
    #spaltenwahl_controller.dispose
    p :closed
  end

  def excel_erstellen_btn_action_performed
    eg = ExcelGenerator.new( model.daten_modell )
  end
end
