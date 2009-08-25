
if not defined? ATS_APP_TEST_DIR_MAPPING then
  
ATS_APP_DIR_NAME = "app" unless defined? ATS_APP_DIR_NAME
ATS_APP_DIR_PREFIX = "ats" unless defined? ATS_APP_DIR_PREFIX

ATS_TEST_DIR_NAME = "test" unless defined? ATS_TEST_DIR_NAME
ATS_TEST_DIR_PREFIX = "ta" unless defined? ATS_TEST_DIR_PREFIX

ATS_APP_TEST_DIR_MAPPING = {"controllers" => "functional", "models" => "unit"}

ATS_BASIS_PFAD = File.expand_path(File.dirname(__FILE__))

def testpfad(app_pfad)
  app_fn = File.expand_path(app_pfad)
  
  tmp_fn = app_fn.sub(/\/#{ATS_APP_DIR_NAME}(\/#{ATS_APP_DIR_PREFIX})?/) do |s| 
        "/#{ATS_TEST_DIR_NAME}#{$1 && '/'+ATS_TEST_DIR_PREFIX}"
  end
  ATS_APP_TEST_DIR_MAPPING.each do |appdir, testdir|  
    tmp_fn.sub!(/(\/#{ATS_TEST_DIR_NAME}\/)#{appdir}(\/|$)/, '\1'+testdir+'\2')
  end
  test_fn = tmp_fn.sub(/\/([^\/]+\.rb)$/, '/test_\1')   
  if not File.exist? test_fn then
    test_fn = test_fn.sub(/(^|\/)test_/, '\1').sub(/(\.rb)$/, '_test\1')
  end
  test_fn
end

def apppfad(test_pfad)
  test_fn = File.expand_path(test_pfad)
  
  tmp_fn = test_fn.sub(/\/#{ATS_TEST_DIR_NAME}(\/#{ATS_TEST_DIR_PREFIX})?/) do |s| 
      "/#{ATS_APP_DIR_NAME}#{$1 && '/'+ATS_APP_DIR_PREFIX}"
  end   
  ATS_APP_TEST_DIR_MAPPING.each do |appdir, testdir|  
    tmp_fn.sub!(/(\/#{ATS_APP_DIR_NAME}\/)#{testdir}(\/|$)/, '\1'+appdir+'\2')
  end  
  app_fn = tmp_fn.sub(/(^|\/)test_/, '\1').sub(/_test(\.[\w]+)$/, '\1')
  app_fn
end

def ats_loadpfade_sicherstellen
  $:.unshift ATS_BASIS_PFAD unless $:.include? ATS_BASIS_PFAD
  $:.unshift testpfad(ATS_BASIS_PFAD) unless $:.include? testpfad(ATS_BASIS_PFAD)  
end

def durchlaufe_unittests(fn)
  ats_loadpfade_sicherstellen
  
  file_symbol = File.basename(fn).gsub( /[^\w\d]/,"_").upcase
  trc_hinweis :utest_fuer, file_symbol if respond_to?(:trc_hinweis)
  
  test_filename = testpfad(fn)
  p [:testpfad, test_filename]
  self.class.module_eval "TESTSTARTED_#{file_symbol}=true; require '#{test_filename}'"
end

else
# TODO #*# Dieser Zweig wird durchlaufen, sollte es aber im Normalfall nicht werden!
#  p "defined ATS_APP_TEST_DIR_MAPPING!!!!!!"
end # if not defined? ATS_APP_TEST_DIR_MAPPING

