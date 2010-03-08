


import org.jruby.Ruby;
import org.jruby.RubyClass;
import org.jruby.runtime.ObjectAllocator;
import org.jruby.runtime.builtin.IRubyObject;
import org.jruby.runtime.load.BasicLibraryService;
import org.jruby.ext.win32ole.WIN32OLE;


public class Win32oleService implements BasicLibraryService {
    public boolean basicLoad(Ruby runtime) {
    	RubyClass win32ole = 
    		runtime.defineClass("WIN32OLE", runtime.getObject(), WIN32OLE_ALLOCATOR);
    	win32ole.defineAnnotatedMethods(WIN32OLE.class);
        return true;
    }
    
    private static ObjectAllocator WIN32OLE_ALLOCATOR = new ObjectAllocator() {
    	public IRubyObject allocate(Ruby runtime, RubyClass klazz) {
    		return new WIN32OLE(runtime, klazz);
    	}
    };
}