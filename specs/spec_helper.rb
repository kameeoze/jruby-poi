require File.expand_path(File.join(File.dirname(__FILE__), '..', 'lib', 'poi'))

class TestDataFile
  def self.expand_path(name)
    File.expand_path(File.join(File.dirname(__FILE__), 'data', name))
  end
end

Spec::Runner.configure do |config|
end
