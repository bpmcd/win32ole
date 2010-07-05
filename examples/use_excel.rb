require 'win32ole'

#
excel = WIN32OLE.new('Excel.Application')

puts "PUBLIC_METHODS"
puts excel.public_methods
puts "PRIVATE_METHODS"
puts excel.private_methods

excel.visible= true
#excel.Quit()
