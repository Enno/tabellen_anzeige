class DateiauswahlView < ApplicationView
  set_java_class 'dateiauswahl.DateiauswahlDialog'

  def create_main_view_component
    DateiauswahlDialog.new(nil, true)
  end

  def load

  end

  map :view => "dateiauswahl_filechooser.selectedFile", :model => :zielpfad, :using => [nil, :string_zu_ruby]

  def string_zu_ruby(jstring)
    jstring.to_s
  end

end
