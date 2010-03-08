package org.jruby.ext.win32ole;


import java.util.Iterator;
import java.util.List;

import org.jruby.Ruby;
import org.jruby.RubyClass;
import org.jruby.RubyHash;
import org.jruby.RubyObject;
import org.jruby.anno.JRubyMethod;
import org.jruby.javasupport.JavaUtil;
import org.jruby.runtime.Block;
import org.jruby.runtime.builtin.IRubyObject;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComException;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.EnumVariant;
import com.jacob.com.SafeArray;
import com.jacob.com.Variant;
import org.jruby.RubyArray;
import org.jruby.anno.JRubyClass;
import org.jruby.runtime.ThreadContext;
import org.jruby.runtime.Visibility;

/**
 * PoC to implement the WIN32OLE extension.
 *
 * This exposes COM objects (that implement IDispatch) to jRuby.
 * Arguments and Return values are casted (copy by value) to
 * apropriate/aproximate jRuby types.
 *
 * The glory details of interfacing with COM are handled by the
 * Jacob (JAva-COM Bridge) library.
 *
 * This is my first jRuby extension!
 *
 * No optmizations done...
 * Not all Variant types are yet handled...
 * ComException is not yet translated to WIN32OLERuntimeError...
 * The "host" argument is not supported by Jacob...
 *
 * @author Rui Lopes (rgl ruilopes com)
 * @author Brian McDevitt
 */
@JRubyClass(name="WIN32OLE", parent="RubyObject")
public class WIN32OLE extends RubyObject {
    private static final long serialVersionUID = 1L;
    protected Dispatch dispatch;

    public WIN32OLE(Ruby runtime, RubyClass type) {
        super(runtime, type);
    }

    public WIN32OLE(Ruby runtime, RubyClass type, Dispatch dispatch) {
        super(runtime, type);
        this.dispatch = dispatch;
    }

    public Dispatch getDispatch() {
        return dispatch;
    }

    @JRubyMethod(name="new", meta=true)
    public static IRubyObject rbNew(IRubyObject cls, IRubyObject id, Block unusedBlock) {
    	ActiveXComponent com = ActiveXComponent.createNewInstance(clsIdOrProgIdFrom(id));
    	if (com == null) {
    		return cls.getRuntime().getNil();
    	}
    	return new WIN32OLE(cls.getRuntime(), (RubyClass) cls, com.getObject());
    }
    
    @JRubyMethod
	public static IRubyObject connect(IRubyObject self, IRubyObject clsIdOrProgId) {
		ActiveXComponent com = ActiveXComponent.connectToActiveInstance(clsIdOrProgIdFrom(clsIdOrProgId));
		if (com == null) {
	    	// XXX connectToActiveInstance does not throw throw an exception... instead it returns null :(
			return self.getRuntime().getNil();
		}
		return new WIN32OLE(self.getRuntime(), self.getType(), com.getObject());
    }

    @JRubyMethod
    public IRubyObject each(ThreadContext cxt, Block block) {
        IRubyObject result = getRuntime().getNil();
        if (block == null || !block.isGiven())
            return result; // XXX raise error instead?
        // TODO if the current object isn't a IEnumVARIANT, we should probably
        //      ignore this call?
        EnumVariant e = new EnumVariant(dispatch);
        try {
            while (e.hasMoreElements()) {
                Variant element = (Variant) e.nextElement();
                IRubyObject value = jacobToRuby(cxt, element);
                // eager release non IDispatch variants
                if (!(value instanceof WIN32OLE))
                    element.safeRelease();
                result = block.yield(cxt, value);
            }
            return result;
        } finally {
        	// eager release the variant because we known we
        	// don't need it anymore.
        	e.safeRelease();
        }
    }

    @JRubyMethod(name = "method_missing", rest = true) 
    public IRubyObject method_missing(ThreadContext context, IRubyObject[] args, Block block) {

        if (args.length < 1)
            throw context.getRuntime().newArgumentError("Too few arguments (0 for 1)");
        // TODO make sure no block was given?

        String name;
        IRubyObject[] newArgs;
        // remove the property/method name from args.
        if ("invoke".equalsIgnoreCase(args[0].asJavaString())) {
        	name = args[1].asJavaString();
        	newArgs = new IRubyObject[args.length - 2];
        	System.arraycopy(args, 2, newArgs, 0, newArgs.length);
        } else {
        	name = args[0].asJavaString();
        	newArgs = new IRubyObject[args.length - 1];
        	System.arraycopy(args, 1, newArgs, 0, newArgs.length);
        }
        args = newArgs;

        try {
            if (name.endsWith("=")) {
                // Property put.
                if (args.length > 1)
                    throw context.getRuntime().newArgumentError("Too many arguments (" + args.length + " for 1)");
                name = name.substring(0, name.length()-1);
                Object value = rubyToJacob(args[0]);
                Dispatch.put(getDispatch(), name, value);
                return context.getRuntime().getNil();
            } else {
                // Property get, or Method call.
                return jacobToRuby(context, Dispatch.callN(getDispatch(), name, rubyToJacob(args)));
            }
        } catch (ComException e) {
            // TODO try to raise a better exception type.
//        	e.printStackTrace();
            RubyArray array = RubyArray.newArrayNoCopy(context.getRuntime(), args);
            throw context.getRuntime().newNoMethodError("No method `" + name + "'", name, array);
        }
    }
    
    @JRubyMethod(name="[]", required = 1)
    public IRubyObject getPropertyWithBrackets(ThreadContext cxt, IRubyObject property) {
    	return jacobToRuby(cxt, Dispatch.call(getDispatch(), "item", property.asJavaString()));
    }
    
    @JRubyMethod(name="[]=", required = 2)
    public IRubyObject setPropertyWithBrackets(ThreadContext cxt, IRubyObject property, IRubyObject newValue) {
    	Object[] args = new Object[] {property.asJavaString(), rubyToJacob(newValue)};
    	Dispatch.invoke(getDispatch(), "item", Dispatch.Put, args, new int[1]);
    	return cxt.getRuntime().getNil();
    }

    private static String clsIdOrProgIdFrom(IRubyObject rubyString) {
		String clsIdOrProgId = rubyString.asString().toString();

        // Jacob accept one of:
        //  * monikor name.  eg: IIS://localhost/w3svc
        //                   eg: clsid:9BA05972-F6A8-11CF-A442-00A0C90A8F39
        //    Internally, it uses CoGetObject.
        //    This is triggered by the presence of a ':' character.
        //  * program id (progid).  eg: InternetExplorer.Application
        //
        // Though, MRI:
        //  * does not accept monikor names
        //  * accepts progid as {9BA05972-F6A8-11CF-A442-00A0C90A8F39}

        // transform names like:
        //   {9BA05972-F6A8-11CF-A442-00A0C90A8F39}
        // into:
        //   clsid:9BA05972-F6A8-11CF-A442-00A0C90A8F39
        if (clsIdOrProgId.startsWith("{"))
        	clsIdOrProgId = "clsid:" + clsIdOrProgId.substring(1, clsIdOrProgId.length()-1);

		return clsIdOrProgId;
	}

    private static Object[] rubyToJacob(IRubyObject[] args) {
   	 	// XXX unfortunately, there is no convertRubyArrayToJava
    	//return JavaUtil.convertJavaArrayToRuby(args);
        Object[] result = new Object[args.length];
        for (int n = 0; n < args.length; ++n)
            result[n] = rubyToJacob(args[n]);
        return result;
    }

    private static Object rubyToJacob(IRubyObject value) {
    	// XXX convertRubyToJava can return null
    	// XXX convertRubyToJava does not support arrays
        if (value instanceof WIN32OLE)
            return ((WIN32OLE) value).getDispatch();
        else if (value instanceof RubyArray) {
        	RubyArray rbAry = (RubyArray) value;        	
        	IRubyObject[] ary = new IRubyObject[rbAry.size()];
        	for (int i = 0; i < rbAry.size(); i++) {
        		ary[i] = (IRubyObject) rbAry.entry(i);
        	}
        	return rubyToJacob(ary);
        } else if (value instanceof RubyHash) {
        	IRubyObject[] ary = new IRubyObject[((RubyHash) value).size()];
        	RubyHash hash = (RubyHash) value;
        	int i = 0;
        	for (Object o : hash.directValues()) {
        		ary[i++] = (IRubyObject) o;
        	}
        	return rubyToJacob(ary);
        }
    	return JavaUtil.convertRubyToJava(value);
    }

    private static IRubyObject jacobToRuby(ThreadContext context, Variant value) {
        if (value == null || value.isNull())
            return context.getRuntime().getNil();

        Object object = value.toJavaObject();
        if (object instanceof Dispatch) 
        	return new WIN32OLE(context.getRuntime(), context.getRuntime().getClass("WIN32OLE"), (Dispatch) object);
        else if (object instanceof SafeArray)
        	return safeArrayToRubyArray(context, (SafeArray)object);
        
        return JavaUtil.convertJavaToRuby(context.getRuntime(), object);
    }
    
    public static RubyArray safeArrayToRubyArray(ThreadContext context, SafeArray sa) {
    	Variant[] variants = sa.toVariantArray();
    	RubyArray ary = RubyArray.newArray(context.getRuntime(), variants.length);
    	for (int i = 0; i < variants.length; i++) {
    		ary.add(jacobToRuby(context, variants[i]));
    	}
    	return ary;
    }
    
    @Override
    protected void finalize() throws Throwable {
    	try {
    		getDispatch().safeRelease();
    		ComThread.Release();
    	} finally {
    		super.finalize();
    	}
    }
}