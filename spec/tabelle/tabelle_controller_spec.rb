# To change this template, choose Tools | Templates
# and open the template in the editor.

p $CLASSPATH

require 'manifest'
p $CLASSPATH

#include_class 'TabelleFrame'
p $LOAD_PATH

require 'tabelle/tabelle_controller'
#require 'rubygems'
#require 'spec'
#require 'spec/mocks/framework'

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
    @tc = TabelleController.instance
  end

  it "should desc" do
    @tc.open
    
  end
end

