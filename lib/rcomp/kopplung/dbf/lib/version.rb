if not defined?(TESTSCHMIEDE_VERSION) then
  module Ats #:nodoc: Zentest SKIP
    module VERSION #:nodoc:
      MAJOR = 0
      MINOR = 9
      TINY  = 9
      ULTRA = 8
      SUBLIME = 2
    end
  end
  
  class Version 
    def initialize(x)
      @varray = case x
        when String then x.scan(/\d+/)
        when Array  then x
        else raise "Kann Version nicht mit #{x.inspect} initialisieren"
      end.map {|y| y.to_i}
    end
    attr_reader :varray
    
    def <=>(other)
      self.varray <=> other.varray
    end
    include Comparable
    
    def to_s
      varray.join(".")
    end
  end

  module Ats #:nodoc: Zentest SKIP
    module VERSION #:nodoc:
      ARRAY  = [MAJOR, MINOR, TINY, ULTRA, SUBLIME]
      STRING = ARRAY.join('.')
      OBJ    = Version.new(ARRAY)
    end
  end
  TESTSCHMIEDE_VERSION = Ats::VERSION::STRING
#  $stderr.puts "neu TESTSCHMIEDE_VERSION"

else
  $stderr.puts "wieder TESTSCHMIEDE_VERSION"
  $stderr.puts caller 
end


