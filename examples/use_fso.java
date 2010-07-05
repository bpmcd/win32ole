import java.util.Enumeration;

import com.jacob.com.Dispatch;
import com.jacob.com.EnumVariant;
import com.jacob.com.Variant;

// javac -cp lib\jacob.jar use_jacob.java
// java -cp lib\jacob.jar;. -Djava.library.path=lib use_jacob
public class use_fso {
	public static void main(String[] args) {
		// NB: You can see the typelibrary using the OleView.exe (C:\Program Files\Microsoft Visual Studio 8\Common7\Tools\Bin)
		//     By selecting the "Type Libraries" and inside it, then "Microsoft Scripting Runtime".
		//     And finally, the "coclass FileSystemObject"

		/*
		// using ActiveXComponent...
		ActiveXComponent fso = ActiveXComponent.createNewInstance("Scripting.FileSystemObject");
		ActiveXComponent folder = new ActiveXComponent(fso.invoke("GetFolder", ".").toDispatch()); // IFolder
		ActiveXComponent files = folder.getPropertyAsComponent("Files"); // IFilesCollection

		for (Enumeration e = new EnumVariant(files.getObject()); e.hasMoreElements(); ) {
			Variant fileVariant = (Variant) e.nextElement();
			ActiveXComponent file = new ActiveXComponent(fileVariant.toDispatch()); // IFile
			System.out.println(file.getPropertyAsString("Path"));
		}
		*/

		// using strait Dispatch...
		Dispatch fso = new Dispatch("Scripting.FileSystemObject");				// IFileSystem3
		Dispatch folder = Dispatch.call(fso, "GetFolder", ".").toDispatch();	// IFolder
		Dispatch files = Dispatch.get(folder, "Files").toDispatch();			// IFilesCollection
		for (Enumeration<Variant> e = new EnumVariant(files); e.hasMoreElements(); ) {
			Variant fileVariant = (Variant) e.nextElement();
			Dispatch file = fileVariant.toDispatch();							// IFile
			System.out.println(Dispatch.get(file, "Path").getString());
		}
	}
}