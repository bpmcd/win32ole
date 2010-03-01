# start me yup with jruby -I..\lib use_ie.rb
# NB: make sure jruby script has java.library.path set, eg: _VM_OPTS=-Djava.library.path=lib;..\lib
require 'win32ole'

# See "How to connect to a running instance of Internet Explorer" at http://support.microsoft.com/kb/176792

puts 'XXX: this example is not done!'
puts 'NB: You need to have IE running for this to work...'

ie = WIN32OLE.connect('InternetExplorer.Application') # XXX we can't connect to IE like this...
ie.Navigate 'http://jruby.org'

puts 'good bye...'