import java.io.IOException;

public class RunTimeExe {
    public static void main(String[] args) {
        Runtime runtime = Runtime.getRuntime();
        String command = "tesseract.exe C:/pdf/image/1.png C:/pdf/image/result -l eng -c tessedit_char_whitelist=0123456789";
        try{
            runtime.exec(command);
        }catch (IOException e){
            e.printStackTrace();
            System.out.println("error: "+e.getMessage());
        }
    }
}