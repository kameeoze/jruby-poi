require File.expand_path(File.join(File.dirname(__FILE__), '..', 'lib', 'poi'))
Dir[File.dirname(__FILE__) + "/support/**/*.rb"].each {|f| require f}

class TestDataFile
  def self.expand_path(name)
    File.expand_path(File.join(File.dirname(__FILE__), 'data', name))
  end
end

Spec::Runner.configure do |config|
end
