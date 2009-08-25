# To change this template, choose Tools | Templates
# and open the template in the editor.
p __FILE__
require 'tabelle/daten_modell'
require '../../cachedcalc/lib/rechen_zeile'

class  RechTabBsp1 < RechenZeile
  belongs_to Fundament
  def_schluessel :schl1
  ausgabe_groeszen :ausg1

  def ausg1
    schl1 + rt1meth(7).to_s * rt1meth(6)
  end

  def rt1meth(arg)
    p [:rt1meth, arg]
    $rt1meth_arg = arg
  end


end

$rt1a = TABELLEN_FUNDAMENT.RechTabBsp1("a")

describe DatenModell do
  before(:each) do
    @dm = DatenModell.new([$rt1a])
  end

  it "should have columns" do
    @dm.getColumnCount.should == 1
    @dm.getColumnName(0).should == "ausg1"
  end
end

