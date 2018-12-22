package mj.hotels;
import mj.hotels.main.MainSwingController;

public class Hotels {
	
	public static boolean DEBUG = true;

	public static void main(String[] args) {
		init();
	}
	
	static void init() {
		// initiate
		MainSwingController controller = new MainSwingController();
	}

	public static void log(String message) {
		if (DEBUG)
            System.out.println(message);
	}

}
