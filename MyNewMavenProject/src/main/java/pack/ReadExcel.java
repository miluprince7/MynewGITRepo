package pack;

import java.io.IOException;

public class ReadExcel {

	public static void main(String[] args) throws IOException {
		Excel e=new Excel();
		String n=e.readData(0,0);
		System.out.println(n);
		String c=e.readData(0,1);
		System.out.println(c);

	}

}
