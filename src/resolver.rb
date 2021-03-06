module Monkeybars
  class Resolver
    IN_FILE_SYSTEM = :in_file_system
    IN_JAR_FILE = :in_jar_file
    
    # Returns a const value indicating if the currently executing code is being run from the file system or from within a jar file.
    def self.run_location
      if File.expand_path(__FILE__) =~ /\.jar\!/
        IN_JAR_FILE
      else
        IN_FILE_SYSTEM
      end
    end
  end
end

class Object
  def add_to_classpath(path)
    $CLASSPATH << get_expanded_path(path)
  end
  
  def add_to_load_path(path)
    $LOAD_PATH << get_expanded_path(path)
  end
  
  def robust_expand_path(path, base_path)
    pure_base_path = base_path
    prefix = pure_base_path.slice!(/^file\:/)
    prefix.to_s + File.expand_path(path.gsub("file:", ""), base_path)
  end


  private
  def get_expanded_path(path)
    #2009-08-13, Sven:
    resolved_path = robust_expand_path(path.gsub("\\", "/"), File.dirname(__FILE__) )
    #orig:
    #resolved_path = File.expand_path(File.dirname(__FILE__) + "/" + path.gsub("\\", "/"))
    resolved_path.gsub!("file:", "") unless resolved_path.index(".jar!")
    resolved_path.gsub!("%20", ' ')
    resolved_path
  end
end
