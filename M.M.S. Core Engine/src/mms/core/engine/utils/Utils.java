package mms.core.engine.utils;

import java.io.File;
import java.io.FilenameFilter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.TimeUnit;

public class Utils {

	public static boolean emptyCache(String templatesFolder) {
		File file = new File(templatesFolder);

		if(file.isDirectory()){
			List<File> listFile = new ArrayList<File>();

			listFile = Arrays.asList(file.listFiles(new FilenameFilter() {
				@Override
				public boolean accept(File dir, String name) {
					return name.toLowerCase().endsWith(".pdf");
				}
			}));
			
			if (listFile.size() == 0) {
				return true;
			} else {
				for(File f : listFile){
					if (!f.delete())
						return false;
				}
				
				return true;
			}
		}

		return false;
	}

	public static String ms2HMS(long ms) {
		return String.format("%02d:%02d:%02d", 
				TimeUnit.MILLISECONDS.toHours(ms),
				TimeUnit.MILLISECONDS.toMinutes(ms) - TimeUnit.HOURS.toMinutes(TimeUnit.MILLISECONDS.toHours(ms)),
				TimeUnit.MILLISECONDS.toSeconds(ms) - TimeUnit.MINUTES.toSeconds(TimeUnit.MILLISECONDS.toMinutes(ms)));
	}

	public static String readableFileSize(long bytes, boolean si) {
		int unit = si ? 1000 : 1024;

		if (bytes < unit) return bytes + "B";

		int    exp = (int)(Math.log(bytes) / Math.log(unit));
		String pre = (si ? "kMGTPE" : "KMGTPE").charAt(exp - 1) + (si ? "" : "i");

		return String.format("%.1f %sB", bytes / Math.pow(unit, exp), pre);
	}

}
