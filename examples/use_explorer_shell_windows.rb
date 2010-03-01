# start me yup with jruby -I..\lib use_explorer_shell_windows.rb
# NB: make sure jruby script has java.library.path set, eg: _VM_OPTS=-Djava.library.path=lib;..\lib
require 'win32ole'

# These clsid came from shdocvw.dll type library (tlb/olb).
CLSID_ShellWindows = '{9BA05972-F6A8-11CF-A442-00A0C90A8F39}' # 'Shell.Application' coclass.

# Lets see which shell windows are opened.
shell = WIN32OLE.new(CLSID_ShellWindows)
shell.each do |window|
  puts window.Path
end

puts 'good bye...'

# XXX this hangs here because no one called ComThread.Release() on the Java Side!