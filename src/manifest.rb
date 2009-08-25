$LOAD_PATH.clear #ensure load path is cleared so system gems and libraries are not used (only project gems/libs)
# Load current and subdirectories in src onto the load path
$LOAD_PATH << File.dirname(__FILE__)

gem_names = ["rspec", "facets", "dbf-1.0.5", "spreadsheet-0.6.4", "ruby-ole-1.2.10"]

Dir.glob(File.expand_path(File.dirname(__FILE__) + "/**/*").gsub('%20', ' ')).each do |directory|
  next if gem_names.any? {|gem_name| directory =~ /\/#{gem_name}\// }
  #next if directory =~ /\/facets\//
  # File.directory? is broken in current JRuby for dirs inside jars
  # http://jira.codehaus.org/browse/JRUBY-2289
  $LOAD_PATH << directory unless directory =~ /\.\w+$/
end
# Some JRuby $LOAD_PATH path bugs to check if you're having trouble:
# http://jira.codehaus.org/browse/JRUBY-2518 - Dir.glob and Dir[] doesn't work
#                                              for starting in a dir in a jar
#                                              (such as Active-Record migrations)
# http://jira.codehaus.org/browse/JRUBY-3247 - Compiled Ruby classes produce
#                                              word substitutes for characters
#                                              like - and . (to minus and dot).
#                                              This is problematic with gems
#                                              like ActiveSupport and Prawn

#===============================================================================
# Monkeybars requires, this pulls in the requisite libraries needed for
# Monkeybars to operate.

require 'resolver'

case Monkeybars::Resolver.run_location
when Monkeybars::Resolver::IN_FILE_SYSTEM
  add_to_classpath '../lib/java/monkeybars-1.0.4.jar'
  add_to_classpath '../package/classes'
  #add_to_classpath '../package/classes/tabelle'
  #add_to_classpath '../src/tabelle'
end

def gem(*args)
  # dummy
end

require 'monkeybars'
require 'application_controller'
require 'application_view'

# End of Monkeybars requires
#===============================================================================
#
# Add your own application-wide libraries below.  To include jars, append to
# $CLASSPATH, or use add_to_classpath, for example:
# 
# $CLASSPATH << File.expand_path(File.dirname(__FILE__) + "/../lib/java/swing-layout-1.0.3.jar")
#
# is equivalent to
#
# add_to_classpath "../lib/java/swing-layout-1.0.3.jar"
#
# There is also a helper for adding to your load path and avoiding issues with file: being
# appended to the load path (useful for JRuby libs that need your jar directory on
# the load path).
#
# add_to_load_path "../lib/java"
#

def robust_expand_path(path, base_path)
  pure_base_path = base_path
  prefix = pure_base_path.slice!(/^file\:/)
  prefix.to_s + File.expand_path(path.gsub("file:", ""), base_path)
end

def add_gem_path(gem_path)
  loadpath_meta_path = File.join( robust_expand_path(gem_path, File.dirname(__FILE__)), "meta", "loadpath")
  p [File.exist?(loadpath_meta_path), loadpath_meta_path]
  relative_lib_paths = if File.exist? loadpath_meta_path then
    File.read(loadpath_meta_path).split("\n")
  else
    ["lib"]
  end
  # relative_lib_paths = %w[lib/core lib/lore lib/more] if gem_path =~ /facets/ then 
  relative_lib_paths.each do |rel_lib_path|
    #add_to_load_path robust_expand_path rel_lib_path, gem_path
    full_gem_path = robust_expand_path(gem_path, File.dirname(__FILE__))
    target_path = robust_expand_path(rel_lib_path, full_gem_path)
    p target_path
    add_to_load_path target_path.sub("file:","")
  end
end

gem_names.each do |gem_name|
  case Monkeybars::Resolver.run_location
  when Monkeybars::Resolver::IN_FILE_SYSTEM
    # Files to be added only when running from the file system go here
    add_gem_path "../lib/ruby/#{gem_name}"
  when Monkeybars::Resolver::IN_JAR_FILE
    # Files to be added only when run from inside a jar file
    add_gem_path "#{gem_name}"
  end
end

rcomp_dirs = %w[ffmath]
rcomp_dirs.each do |name|
  add_to_load_path case Monkeybars::Resolver.run_location
  when Monkeybars::Resolver::IN_FILE_SYSTEM
    "../lib/rcomp/#{name}/lib"
  when Monkeybars::Resolver::IN_JAR_FILE
    "#{name}/lib"
  end
end

case Monkeybars::Resolver.run_location
when Monkeybars::Resolver::IN_FILE_SYSTEM
  # Files to be added only when running from the file system go here
when Monkeybars::Resolver::IN_JAR_FILE
  # Files to be added only when run from inside a jar file
end
