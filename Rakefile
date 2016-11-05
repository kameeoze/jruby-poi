require 'bundler/setup'
require 'rake'
require 'rake/clean'
require 'rspec/core/rake_task'
require_relative './lib/poi/version'

NAME = 'jruby-poi'
VERSION = POI.version

desc "Build gem"
task :package=>[:clean] do |p|
  sh %{#{FileUtils::RUBY} -S gem build jruby-poi.gemspec}
end

desc "Publish gem"
task :release=>[:package] do
  sh %{#{FileUtils::RUBY} -S gem push ./#{NAME}-#{VERSION}.gem}
end

desc "Run all tests"
RSpec::Core::RakeTask.new('spec') 

desc "Print version"
task :version do
  puts VERSION
end

task :default => :spec
