# To change this template, choose Tools | Templates
# and open the template in the editor.

p $CLASSPATH

require 'manifest'
p $CLASSPATH

#include_class 'TabelleFrame'
p $CLASSPATH

require 'tabelle/tabelle_controller'
require 'spec/mocks/framework'

class DummyViewComponent
  def method_missing(*args, &blk)
    p [:VIEW, args, blk]
  end
end



class TabelleController
  #set_view 'DummyViewComponent'
end


describe TabelleController do
  before(:each) do
    @tabelle_controller = TabelleController.instance
  end

  it "should desc" do
    # TODO
  end
end

