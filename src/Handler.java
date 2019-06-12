public class Handler {
    public static void main(String[] args) {
        Runtime runtime = Runtime.getRuntime();
        try {
          runtime.exec()
        }
        catch (Exception e) {
          e.printStackTrace();
        }
    }
}
