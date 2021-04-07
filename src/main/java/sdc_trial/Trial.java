package sdc_trial;

public class Trial {

	public static void main(String[] args) {
		if(args.length != 1) {
			System.out.println("Need exactly one argument - the path to an Excel file.");
			System.exit(0);
		}

		ExcelReader.findPrimes(args[0]);
	}

}
