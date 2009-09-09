#import java.io.File

class DateiauswahlView < ApplicationView
  set_java_class 'dateiauswahl.DateiauswahlDialog'

  def create_main_view_component
    DateiauswahlDialog.new(nil, true)
  end

  def load
    default_directory = File.dirname(File.dirname(File.dirname(__FILE__))) + "/daten/"
    dateiauswahl_filechooser.currentDirectory = java.io.File.new(default_directory)
  end

  map :view => "dateiauswahl_filechooser.selectedFile", :model => :destination_path, :using => [nil, :string_zu_ruby]
  
  def string_zu_ruby(jstring)
    jstring.to_s
  end

end
