require 'rubygems'
require 'rake'
require 'rake/clean'
require 'rake/testtask'
require 'rake/packagetask'
require 'rake/gempackagetask'
require 'rake/rdoctask'
require 'rake/contrib/rubyforgepublisher'
require 'fileutils'
require 'hoe'
include FileUtils
require File.join(File.dirname(__FILE__), 'lib', 'auto_office', 'version')

AUTHOR = "Nathan Smith"  # can also be an array of Authors
EMAIL = "nlsmith@nlsmith.com"
DESCRIPTION = "Library for easy automation of office applications"
GEM_NAME = "auto_office" # what ppl will type to install your gem
RUBYFORGE_PROJECT = "auto_office" # The unix name for your project
HOMEPATH = "http://#{RUBYFORGE_PROJECT}.rubyforge.org"
RELEASE_TYPES = %w( gem ) # can use: gem, zip

NAME = "auto_office"
REV = nil # UNCOMMENT IF REQUIRED: File.read(".svn/entries")[/committed-rev="(d+)"/, 1] rescue nil
VERS = ENV['VERSION'] || (AutoOffice::VERSION::STRING + (REV ? ".#{REV}" : ""))
CLEAN.include ['**/.*.sw?', '*.gem', '.config']
RDOC_OPTS = ['--quiet', '--title', "auto_office documentation",
    "--opname", "index.html",
    "--line-numbers", 
    "--main", "README",
    "--inline-source"]

# Generate all the Rake tasks
# Run 'rake -T' to see list of generated tasks (from gem root directory)
hoe = Hoe.new(GEM_NAME, VERS) do |p|
  p.author = AUTHOR 
  p.description = DESCRIPTION
  p.email = EMAIL
  p.summary = DESCRIPTION
  p.url = HOMEPATH
  p.rubyforge_name = RUBYFORGE_PROJECT if RUBYFORGE_PROJECT
  p.test_globs = ["test/**/*_test.rb"]
  p.clean_globs = CLEAN  #An array of file patterns to delete on clean.
  
  # == Optional
  #p.changes        - A description of the release's latest changes.
  p.extra_deps = ['hoe', 'rubyforge']    
  #p.spec_extras    - A hash of extra values to set in the gemspec.
end
          
