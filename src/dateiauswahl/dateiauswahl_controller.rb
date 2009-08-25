
class DateiauswahlController < ApplicationController
  set_model 'DateiauswahlModel'
  set_view 'DateiauswahlView'
  set_close_action :hide

  def dateiauswahl_filechooser_action_performed
    update_model view_model, :destination_path
    hide
  end

  def get_destination_path
    update_model view_model, :destination_path 
    model.destination_path
  end
end
