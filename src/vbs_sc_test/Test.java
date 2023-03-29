package vbs_sc_test;

import com.github.jbrienen.vbs_sc.ShortcutFactory;

import java.io.FileNotFoundException;

public class Test {

	public static void main(String[] args) {
		defaultTest();
		iconTest();
	}

	public static void defaultTest(){
		try {
			ShortcutFactory.createDesktopShortcut("C:/Program Files/Internet Explorer/iexplore.exe", "Internet_Explorer.lnk");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}

	public static void iconTest(){
		try {
			// Uses the icon file found in the res directory under the name 'icon.ico'
			ShortcutFactory.createDesktopShortcut("C:/Program Files/Internet Explorer/iexplore.exe",
												"Internet_Explorer_With_Icon.lnk",
														"res/icon.ico");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}


}
