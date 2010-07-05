# start me yup with jruby -I..\lib use_ie.rb
# NB: make sure jruby script has java.library.path set, eg: _VM_OPTS=-Djava.library.path=lib;..\lib
require 'win32ole'

# See "How to connect to a running instance of Internet Explorer" at http://support.microsoft.com/kb/176792

# Lets create a new Windows Explorer browser Window.
CLSID_ShellBrowser = '{c08afd90-f2a1-11d1-8455-00a0c91f3880}' # Default interface is IWebBrowser2
browser = WIN32OLE.new(CLSID_ShellBrowser)
browser.Visible = true
#browser.Navigate('jruby.org') # oddly enough, this open the URL in the default user browser (in my case, Firefox).
browser.Navigate('file://c:/')
puts "Name=    #{browser.Name}"
puts "FullName=#{browser.FullName}"
puts "Path=    #{browser.Path}"

browser.Quit # NB: for now on, the browser MIGHT be in a fubar state, because it can be disconnected from the Windows Explorer instance.
puts "Name=    #{browser.Name}" # MAYBE this raises exception.

puts 'good bye...'