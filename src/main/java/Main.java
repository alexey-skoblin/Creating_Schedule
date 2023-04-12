import lombok.extern.slf4j.Slf4j;

import java.io.FileNotFoundException;
import java.nio.file.Path;
import java.nio.file.Paths;

@Slf4j
public class Main {
    public static void main(String[] args) {


        log.atInfo().log("Start program");
        Path fileLocation = Paths.get("Расписание.txt");
        try {
            Schedule schedule = ReadFile.readFile(fileLocation);
            if (schedule != null)
                new EditorExcel().writeFile(schedule);
        }
        catch (FileNotFoundException fileNotFoundException)
        {
            log.atError().log("File not found for path: " + fileLocation);
        }
        log.atInfo().log("Exit Program.");
    }
}
