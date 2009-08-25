
class DateiauswahlController < ApplicationController
  set_model 'DateiauswahlModel'
  set_view 'DateiauswahlView'
  set_close_action :hide

  def dateiauswahl_filechooser_action_performed
    update_model view_model, :zielpfad
    hide
  end

  def uebergebe_zielpfad
    update_model view_model, :zielpfad
    puts "zielpfad von der methode : #{model.zielpfad}"
    model.zielpfad
  end
end

#TODO: currentdirectory dynamisch