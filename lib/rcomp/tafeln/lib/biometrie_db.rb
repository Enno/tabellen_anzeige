# To change this template, choose Tools | Templates
# and open the template in the editor.

class BiometrieDb
  attr_reader :nx_quellen
  def initialize
    @nx_quellen = Hash.new
  end
end

BIOMETRIE_DB = BiometrieDb.new