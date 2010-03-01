import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.EnumVariant;
import com.jacob.com.Variant;

// javac -cp lib\jacob.jar use_explorer_shell_windows.java
// java -cp lib\jacob.jar;. -Djava.library.path=lib use_explorer_shell_windows
public class use_explorer_shell_windows {
	public static void main(String[] args) {
		ComThread.InitMTA();
		try {
			Dispatch shellWindows = new Dispatch("clsid:9BA05972-F6A8-11CF-A442-00A0C90A8F39");
			for (EnumVariant e = new EnumVariant(shellWindows); e.hasMoreElements(); ) {
				Variant fileVariant = (Variant) e.nextElement();
				Dispatch file = fileVariant.toDispatch();
				System.out.println(Dispatch.get(file, "Path").getString());
			}
		} finally {
			ComThread.Release();
		}
	}
}