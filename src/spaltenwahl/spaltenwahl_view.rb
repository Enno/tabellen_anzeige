
include_class 'spaltenwahl.SpaltenwahlDialog'

class DummyViewComponent
  def method_missing(*args, &blk)
    p [:VIEW, args, blk]
  end
end

class SpaltenwahlView < ApplicationView
  #set_java_class "spaltenwahl.SpaltenwahlFrame"

  def create_main_view_component
    DummyViewComponent.new
  end

# Hab es jetzt mit Dialogs und nested components versucht:
# http://kenai.com/projects/monkeybars/pages/Dialogs/text
# http://groups.google.com/group/monkeybars/browse_thread/thread/6f0c5e59110b696b (Variante c)
# Funktioniert, bis auf das Abfangen der Ereignisse zum schließen.
# Da das aber wichtig ist, bin ich etwas ratlos.
#
  # statt dessen:
  def parent_component=(parent)
    p :parent_component=
   # ok to instantiate our JDialog now since we know what parent to  attach it to
   @main_view_component = SpaltenwahlDialog.new(parent, true)
  end

  map :view => "spaltenliste", :model => :auswahl_werte, :using => [nil, :hole_werte]

  def hole_werte(jlist_objekt)
    erg = jlist_objekt.getSelectedValues
    p erg
    erg
  end

end
