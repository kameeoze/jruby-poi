require 'rspec'
require File.expand_path(File.join(File.dirname(__FILE__), '..', 'lib', 'poi'))
Dir[File.dirname(__FILE__) + "/support/**/*.rb"].each {|f| require f}
require File.join(File.dirname(__FILE__), "support", "java", "support.jar")

class TestDataFile
  def self.expand_path(name)
    File.expand_path(File.join(File.dirname(__FILE__), 'data', name))
  end
end

RSpec.configure do |config|
end
