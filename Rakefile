begin
  require 'jeweler'
  
  Jeweler::Tasks.new do |gemspec|
    gemspec.name = "jruby-poi"
    gemspec.summary = "Apache POI class library for jruby"
    gemspec.description = "A rubyesque library for manipulating spreadsheets and other document types for jruby, using Apache POI."
    gemspec.email = ["sdeming@makefile.com", "jacaetevha@gmail.com"]
    gemspec.homepage = "http://github.com/sdeming/jruby-poi"
    gemspec.authors = ["Scott Deming", "Jason Rogers"]
  end
rescue LoadError
  puts "Jeweler not available. Install it with: gem install jeweler"
end

begin
  require 'rspec/core/rake_task'
  task :default => :spec

  desc "Run all examples"
  RSpec::Core::RakeTask.new do |t|
    t.pattern = 'specs/**/*.rb'
    t.rspec_opts = ['-c']
  end

  desc "Run all examples with RCov"
  RSpec::Core::RakeTask.new(:coverage) do |t|
    t.pattern = 'specs/**/*.rb'
    t.rspec_opts = ['-c']
    t.rcov = true
    t.rcov_opts = ['--include', '/lib/', '--exclude', '/gems/,/^specs/,/^.eval.$/']
  end
rescue
  puts $!.message
  puts "RCov not available. Install it with: gem install rcov (--no-rdoc --no-ri)"
end
