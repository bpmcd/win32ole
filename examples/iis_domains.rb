# start me yup with jruby -I..\lib iis_domains.rb
# NB: make sure jruby script has java.library.path set, eg: _VM_OPTS=-Djava.library.path=lib;..\lib
#
# An example that displays all the configured domains inside IIS 6+.
# -- Rui Lopes (rgl ruilopes com)
require 'win32ole'

# returns a list with all IIS configured domains names and corresponding
# root path, and log path.
def iis_domains
  domains = []

  # Use the ADSI interface to get all configured domains inside IIS.
  service = WIN32OLE.connect("IIS://localhost/w3svc")
  service.each do |server|
    next unless server.Class == "IIsWebServer"

    #domain = server.ServerComment
    ip, port, domain = server.ServerBindings.first.split(":", 3)

    # ignore servers without a "host" header (except for the default server).
    if domain.empty?
      next unless server.Name == "1"
      domain = "localhost"
    end

    path = server.GetObject("IIsWebVirtualDir", "root").Path
    # NB: Since we don't have the WIN32 extension, I'll not expand the
    #     directory, though, we could use the "Scripting.Shell" COM object...
    #log_path = Win32::Registry.expand_environ(server.LogFileDirectory) + "\\W3SVC" + server.Name
    log_path = server.LogFileDirectory + "\\W3SVC" + server.Name

    domains << [domain, path, log_path]
  end

  domains
end

puts iis_domains
puts 'good bye...'