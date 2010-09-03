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

require 'spec/rake/spectask'

desc "Run all examples with RCov"
Spec::Rake::SpecTask.new('rcov') do |t|
  t.spec_files = FileList['specs/**/*.rb']
  t.spec_opts = ['-c']
  t.rcov = true
  t.rcov_opts = ['--include', '/lib/', '--exclude', '/gems/,/^specs/,/^.eval.$/']
end