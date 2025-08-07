package mms.core.engine.pdfmerger.commons;

public class CommonUtils {

	public static double[] getDoubleArray(String[] strArray) {
		double doubleArray[] = new double[strArray.length];
		
		for (int i = 0; i < strArray.length; i++) 
			doubleArray[i] = Double.valueOf(strArray[i].trim());
	
		return doubleArray;
	}

	public static float[] getFloatArray(String[] strArray) {
		float floatArray[] = new float[strArray.length];
		
		for (int i = 0; i < strArray.length; i++) 
			floatArray[i] = Float.valueOf(strArray[i].trim());
	
		return floatArray;
	}

	public static int[] getIntegerArray(String[] strArray) {
		int    intArray[] = new int[strArray.length];
		
		for (int i = 0; i < strArray.length; i++) 
			intArray[i] = Integer.valueOf(strArray[i].trim());
	
		return intArray;
	}

}
