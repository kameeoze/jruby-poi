require 'bundler/setup'
require 'rspec/core/rake_task'

begin
  require 'jeweler'
  
  Jeweler::Tasks.new do |gemspec|
    gemspec.name = "jruby-poi"
    gemspec.summary = "Apache POI class library for jruby"
    gemspec.description = "A rubyesque library for manipulating spreadsheets and other document types for jruby, using Apache POI."
    gemspec.email = ["sdeming@makefile.com", "jacaetevha@gmail.com"]
    gemspec.homepage = "http://github.com/kameeoze/jruby-poi"
    gemspec.authors = ["Scott Deming", "Jason Rogers"]
  end
rescue LoadError
  puts "Jeweler not available. Install it with: gem install jeweler"
end

begin
  task :default => :spec

  desc "Run all examples"
  RSpec::Core::RakeTask.new(:spec)

  desc "Run all examples with rcov"
  RSpec::Core::RakeTask.new(:coverage) do |t|
    t.rcov = true
    t.rcov_opts = ['--include', '/lib/', '--exclude', '/gems/,/^specs/,/^.eval.$/']
  end
rescue
  puts $!.message
  puts "RCov not available. Install it with: gem install rcov (--no-rdoc --no-ri)"
end
