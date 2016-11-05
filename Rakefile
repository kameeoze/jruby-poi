require 'bundler/setup'
require 'rspec/core/rake_task'

require_relative './lib/poi/version'

NAME = 'jruby-poi'
VERSION = POI.version

desc "Run all tests"
RSpec::Core::RakeTask.new('spec') 

desc "Print version"
task :version do
  puts VERSION
end

task :default => :spec
