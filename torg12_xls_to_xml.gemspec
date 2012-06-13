# -*- encoding: utf-8 -*-
require File.expand_path('../lib/torg12_xls_to_xml/version', __FILE__)

Gem::Specification.new do |gem|
  gem.authors       = ["ror"]
  gem.email         = ["mister-al@ya.ru"]
  gem.description   = %q{Torg12 xls to xml}
  gem.summary       = %q{torg12 xls to xml}
  gem.homepage      = ""

  gem.files         = `git ls-files`.split($\)
  gem.executables   = gem.files.grep(%r{^bin/}).map{ |f| File.basename(f) }
  gem.test_files    = gem.files.grep(%r{^(test|spec|features)/})
  gem.name          = "torg12_xls_to_xml"
  gem.require_paths = ["lib"]
  gem.version       = Torg12XlsToXml::VERSION

  #gem 'spreadsheet'
  #gem 'nokogiri'
  gem.add_dependency('nokogiri','>= 1.5.2')
  gem.add_dependency('spreadsheet','>= 0.7.1')
end
