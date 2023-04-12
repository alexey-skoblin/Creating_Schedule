import lombok.extern.slf4j.Slf4j;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

@Slf4j
public class ReadFile {

    static Schedule readFile(Path path) throws FileNotFoundException {
        List<String> lines = new ArrayList<>();
        try {
            lines = Files.readAllLines(path, StandardCharsets.UTF_8);
        } catch (IOException e) {
            log.atError().log("Error read file!");
        }

        try {
            Schedule schedule = new Schedule(lines.get(0));

            for (int i = 1; i < lines.size(); i++) {
                String line = lines.get(i);
                String[] array =  line.split(",");
                if (array.length < 7)
                    continue;

                for(int j = 0; j < array.length; j++)
                {
                    String s = array[j];
                    if (s.startsWith(" "))
                        array[j] = s.substring(1);
                    if (s.endsWith(" "))
                        array[j] = s.substring(0, s.length() -1 );
                }

                schedule.addClass(
                        Integer.parseInt(array[0]),
                        Integer.parseInt(array[1]),
                        EducationClass.convertStringToRotationWeek(array[2]),
                        array[3],
                        EducationClass.convertStringToTypeSubject(array[4]),
                        array[5],
                        array[6].substring(0, array[6].indexOf("/")),
                        EducationClass.convertStringToStudyCorp(array[6].substring(array[6].indexOf("/") + 1))
                );
            }

            for(Schedule.Day day : schedule.getWeek())
            {
                day.sort();
            }

            return schedule;
        } catch (Exception e) {
            e.printStackTrace();
            log.atError().log("Error convert schedule");
            return null;
        }

    }

}
