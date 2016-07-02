import java.io.IOException;

public class ExcelMerger {

	public static void main(String[] args) {
		System.out.println(args[0]);
		System.out.println(args[1]);
		HorizonMerger merger=new HorizonMerger(args[0],args[1]);
		try {
			merger.merge();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
