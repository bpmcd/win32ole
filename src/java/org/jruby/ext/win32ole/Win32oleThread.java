/**
 * 
 */
package org.jruby.ext.win32ole;

import org.jruby.Ruby;
import org.jruby.RubyClass;
import org.jruby.RubyModule;
import org.jruby.anno.JRubyMethod;
import org.jruby.anno.JRubyModule;
import org.jruby.runtime.Block;
import org.jruby.runtime.ThreadContext;
import org.jruby.runtime.builtin.IRubyObject;

import com.jacob.com.ComThread;

/**
 * @author bpm
 *
 */
@JRubyModule(name="ComThread")
public class Win32oleThread extends RubyModule {
	
	
    /**
	 * 
	 */
	private static final long serialVersionUID = -5478671949129473081L;

	protected Win32oleThread(Ruby runtime) {
		super(runtime);
	}

	public Win32oleThread(Ruby runtime, RubyClass klazz) {
		super(runtime, klazz);
	}

	private static class IntThreadLocal extends ThreadLocal<Integer> {
		protected Integer initialValue() {
		    return Integer.valueOf(0);
		}
    }

    private static ThreadLocal<Integer> countMTA = new IntThreadLocal();
    private static ThreadLocal<Integer> countSTA = new IntThreadLocal();

    @JRubyMethod(module=true)
    public static void inApartment(ThreadContext context, IRubyObject unused, Block block) {
	if (((Integer)countSTA.get()).intValue() > 0)
	    withSTA(context, unused, block);
	else
	    withMTA(context, unused, block);
    }

    @JRubyMethod(module=true)
    public static void withSTA(ThreadContext context, IRubyObject unused, Block block) {
		initSTA(context);
		try {
		    block.yield(context, null);
		} finally {
		    releaseSTA(context);
		}
    }

    @JRubyMethod(module=true)
    public static void withMTA(ThreadContext context, IRubyObject unused, Block block) {
		initMTA(context);
		try {
		    block.yield(context, null);
		} finally  {
		    releaseMTA(context);
		}
    }

    synchronized static void initSTA(ThreadContext context) {
		if(((Integer)countMTA.get()).intValue() > 0)
		    throw context.getRuntime().newRuntimeError("Cannot initialize STA thread; current thread is MTA.");
	
		if(((Integer)countSTA.get()).intValue() == 0)
		    ComThread.InitSTA();
	
		countSTA.set(Integer.valueOf(((Integer)countSTA.get()).intValue() + 1));
    }
    
    synchronized static void releaseSTA(ThreadContext context) {
		if(((Integer)countSTA.get()).intValue() == 0)
		    throw context.getRuntime().newRuntimeError("Current thread is not STA.");
	
		countSTA.set(Integer.valueOf(((Integer)countSTA.get()).intValue() - 1));
		
		if(((Integer)countSTA.get()).intValue() == 0)
		    ComThread.Release();
    }

    public synchronized static void initMTA(ThreadContext context) {
		if(((Integer)countSTA.get()).intValue() > 0)
		    throw context.getRuntime().newRuntimeError("Cannot initialize MTA thread; current thread is STA.");
	
		if (((Integer)countMTA.get()).intValue() == 0)
		    ComThread.InitMTA();
	
		countMTA.set(Integer.valueOf(((Integer)countMTA.get()).intValue() + 1));
    }

    public synchronized static void releaseMTA(ThreadContext context) {
		if(((Integer)countMTA.get()).intValue() == 0)
		    throw context.getRuntime().newRuntimeError("Current thread is not MTA.");
	
		countMTA.set(Integer.valueOf(((Integer)countMTA.get()).intValue() - 1));
	
		if(((Integer)countMTA.get()).intValue() == 0)
		    ComThread.Release();
    }

}
