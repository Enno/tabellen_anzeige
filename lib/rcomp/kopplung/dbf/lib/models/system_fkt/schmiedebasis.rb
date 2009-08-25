
# Diese Datei bewirkt nur ein sicheres:
#            require "../schmiedebasis"

pfad_teile = File.expand_path(__FILE__).split("/")
pfad_teile.slice!(-2)
require pfad_teile.join("/")

