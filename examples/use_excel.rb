require 'win32ole'

#
excel = WIN32OLE.new('Excel.Application')
##debugger
puts excel.object_id
excel.Visible = true
excel.Quit
