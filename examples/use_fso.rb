# start me yup with jruby -I..\lib use_fso.rb
# NB: make sure jruby script has java.library.path set, eg: _VM_OPTS=-Djava.library.path=lib;..\lib
require 'win32ole'

# NB: You can see the typelibrary using the OleView.exe (C:\Program Files\Microsoft Visual Studio 8\Common7\Tools\Bin)
#     By selecting the "Type Libraries" and inside it, then "Microsoft Scripting Runtime".
#     And finally, the "coclass FileSystemObject"
fso = WIN32OLE.new('Scripting.FileSystemObject')
folder = fso.GetFolder('.') # XXX if use a non-existent folder, the error is cryptic :|
folder.SubFolders.each do |file|
    puts file.Path + '\\'
end
folder.Files.each do |file|
    puts file.Path
end

puts 'good bye...'