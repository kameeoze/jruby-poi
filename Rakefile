require 'rubygems'

require 'bundler'
begin
  Bundler.setup(:default, :development)
rescue Bundler::BundlerError => e
  $stderr.puts e.message
  $stderr.puts "Run `bundle install` to install missing gems"
  exit e.status_code
end

require 'rake'
require 'jeweler'
require 'rspec/core'
require 'rspec/core/rake_task'

# gemify
Jeweler::Tasks.new do |gemspec|
  gemspec.name = "jruby-poi"
  gemspec.summary = "Apache POI class library for jruby"
  gemspec.description = "A rubyesque library for manipulating spreadsheets and other document types for jruby, using Apache POI."
  gemspec.license = "Apache"
  gemspec.email = ["sdeming@makefile.com", "jacaetevha@gmail.com"]
  gemspec.homepage = "http://github.com/kameeoze/jruby-poi"
  gemspec.authors = ["Scott Deming", "Jason Rogers"]
end
Jeweler::RubygemsDotOrgTasks.new

# test
RSpec::Core::RakeTask.new(:spec)
RSpec::Core::RakeTask.new(:rcov) do |spec|
  spec.rcov = true
end

task :default => :spec
