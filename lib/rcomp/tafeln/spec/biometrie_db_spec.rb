# encoding: utf-8


require 'spec'

require File.expand_path(__FILE__).sub("_spec.rb", ".rb").sub("spec/", "lib/")
require File.dirname(File.dirname(File.expand_path(__FILE__))) + "/lib/db_excel"
#require 'biometrie_db'require File.

describe BiometrieDb do
  before(:all) do
    @bdb = BIOMETRIE_DB #BiometrieDb.new
  end

  it "should desc" do
    nx_quelle = @bdb.nx_quellen["F1994T"]
    nx_quelle.zins.should == 0.04
    nx_quelle.art.should == nil
    nx_quelle.nx(0).should == 0
    nx_quelle.nx(1).should == 23290350.365
  end
end

