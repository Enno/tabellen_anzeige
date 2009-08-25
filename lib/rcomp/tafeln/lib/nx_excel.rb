# encoding: utf-8

class NxExcel
  attr_reader :zins, :art
  def initialize(spalte)
    @zins, @art, dummy, *@nx = spalte
    #@nx.reverse_each {|x| @nx}
  end

  def nx(x)
    @nx[x]
  end
end
