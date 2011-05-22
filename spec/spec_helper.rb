$stderr.puts "included spec_helper.rb"
RSpec.configure do |c|
  c.filter_run_excluding :unimplemented => true
end

require File.expand_path('../lib/poi', File.dirname(__FILE__))

Dir[File.dirname(__FILE__) + "/support/**/*.rb"].each {|f| require f}
require File.join(File.dirname(__FILE__), "support", "java", "support.jar")

class TestDataFile
  def self.expand_path(name)
    File.expand_path(File.join(File.dirname(__FILE__), 'data', name))
  end
end
