require 'ffi'
require 'win32ole.jar'

class WIN32OLE
  extend FFI::Library
  ffi_lib 'kernel32'
  ffi_convention :stdcall

  CP_UTF8       = 65001
  CP_UTF7       = 65000
  CP_ACP        = 0
  CP_OEMCP      = 1
  CP_MACCP      = 2
  CP_THREAD_ACP = 3
  CP_SYMBOL     = 42
  LOCALE_SYSTEM_DEFAULT = 0x0800
  LOCALE_USER_DEFAULT   = 0x0400

  def added
    nil
  end
  
  attach_function :GetACP, [], :uint
 # attach_function :SetCodePage, [:long, :long], :uint

  def codepage
    self.GetACP
  end

#  def codepage=(cp)
#    self.SetCodePage(cp)
#    nil
#  end
end
