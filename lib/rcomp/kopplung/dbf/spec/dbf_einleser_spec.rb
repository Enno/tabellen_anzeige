# encoding: utf-8

require File.expand_path(__FILE__).sub("_spec.rb", ".rb").sub("spec/", "lib/")
#require 'dbf_einleser'

describe DbfEinleser do
  before(:all) do
    # @de = DbfEinleser.new("/dat/GiS/gm/MStar/MsNeu/DATEN/BAT7V2")
    $db_pfad = "/dat/GiS/gm/MStar/MsNeu/DATEN/AT7V2"
    @de = DbfEinleser.new("komp", :vk)
  end

  it "should " do
    @de.alle_schluessel.size.should == 8
    @de.lese("A1").komp.should == "A1"
  end

  it "should simply read dbf file" do
    @de.should_not be_nil
    DbfDat::oeffnen("/dat/GiS/gm/MStar/MsNeu/DATEN/AT7V2", :b  )
    vd = DbfDat::Vd.find( :first, :vsnr => 'Test01' )
    vd.should_not be_nil
    #p vd.size
    vd.vsnr.should == "Test01"
  end
end

