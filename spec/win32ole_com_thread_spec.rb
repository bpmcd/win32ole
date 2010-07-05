# win32ole_com_thread_spec.rb
require 'win32ole'

describe ComThread, "#inApartment" do
  it "raises an error when using inApartment before STA" do
    lambda {
      ComThread.inApartment() {
        x = WIN32OLE.new("InternetExplorer.Application")
      	ComThread.withSTA() {
          x.visible = true
          x.gohome
      	}
        x.quit
      }
    }.should raise_error(RuntimeError)
  end

  it "should not raise an error when using STA before inApartment" do
    lambda {
      ComThread.withSTA() {
        x = WIN32OLE.new("InternetExplorer.Application")
      	ComThread.inApartment() {
          x.visible = true
          x.gohome
      	}
        x.quit
      }
    }.should_not raise_error
  end



  it "should not raise an error when using inApartment before MTA" do
    lambda {
      ComThread.inApartment() {
        x = WIN32OLE.new("InternetExplorer.Application")
      	ComThread.withMTA() {
          x.visible = true
          x.gohome
      	}
        x.quit
      }
    }.should_not raise_error
  end

  it "should not raise an error when using MTA before inApartment" do
    lambda {
      ComThread.withMTA() {
        x = WIN32OLE.new("InternetExplorer.Application")
      	ComThread.inApartment() {
          x.visible = true
          x.gohome
      	}
        x.quit
      }
    }.should_not raise_error
  end
end

describe ComThread, "mixing withSTA and withMTA" do
  it "should raise an error when mixing STA and MTA methods" do
    lambda {
      ComThread.withSTA {
        x = WIN32OLE.new("InternetExplorer.Application")
        ComThread.withMTA {
          x.visible = true
          x.gohome
        }
        x.quit
      }
    }.should raise_error(RuntimeError)
  end

  it "should raise an error when mixing STA and MTA methods" do
    lambda {
      ComThread.withMTA {
        x = WIN32OLE.new("InternetExplorer.Application")
        ComThread.withSTA {
          x.visible = true
          x.gohome
        }
        x.quit
      }
    }.should raise_error(RuntimeError)
  end
end
