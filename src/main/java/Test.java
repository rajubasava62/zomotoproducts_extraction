import java.util.*;

class Test{
    public static void main(String[] args) {
        Test t = new Test();
        System.out.println(t.generateId());
    }
    public String generateId() {
        Random random = new Random();
        String digits = "0123456789";
        String letters = "abcdefghijklmnopqrstuvwxyz";
        StringBuilder id = new StringBuilder();

        // Generate 5 digits first
        for (int i = 0; i < 5; i++) {
            int index = random.nextInt(digits.length());
            id.append(digits.charAt(index));
        }

        // Generate 5 letters
        for (int i = 0; i < 5; i++) {
            int index = random.nextInt(letters.length());
            id.append(letters.charAt(index));
        }

        // Shuffle the last 9 characters (to mix numbers and letters after the starting digit)
        char[] idArray = id.substring(1).toCharArray();
        for (int i = idArray.length - 1; i > 0; i--) {
            int j = random.nextInt(i + 1);
            char temp = idArray[i];
            idArray[i] = idArray[j];
            idArray[j] = temp;
        }

        // Append the shuffled part back to the first digit
        return id.charAt(0) + new String(idArray);
    }
}