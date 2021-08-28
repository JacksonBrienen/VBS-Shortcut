package vbs_sc_test;

import java.io.FileNotFoundException;

import vbs_sc.ShortcutFactory;

public class Test {

	public static void main(String[] args) {
		try {
			ShortcutFactory.createDesktopShortcut("C:/Program Files/Internet Explorer/iexplore.exe", "Internet_Explorer.lnk");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}
	
}
