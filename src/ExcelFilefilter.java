import java.io.File;
import java.io.FileFilter;

public class ExcelFilefilter implements FileFilter {

	@Override
	public boolean accept(File file) {
		String name=file.getName();
		String extension=name.substring(name.lastIndexOf(".")+1);
		return "xlsx".equals(extension);
	}

}
