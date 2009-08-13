class SpaltenwahlController < ApplicationController
  set_model 'SpaltenwahlModel'
  set_view 'SpaltenwahlView'
  set_close_action :hide

  #add_listener :type => :action, :components => [:ok_btn]


  def ergebnis
    # hackige Methode, mit der man vielleicht Daten aus dem Dialog ziehen könnte:
    p @__view.instance_variable_get(:@main_view_component)
    #p @__view.spaltenliste
    [:ergebnis]
  end
end
