import lombok.*;
import lombok.extern.slf4j.Slf4j;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

@Slf4j
@Getter
@Setter
@RequiredArgsConstructor
public class Schedule {
    @NonNull
    private String name;
    private Day[] week = {new Day("Понедельник"), new Day("Вторник"), new Day("Среда"), new Day("Четверг"), new Day("Пятница"), new Day("Суббота")};

    @Setter
    @Getter
    @ToString
    @RequiredArgsConstructor
    static class Day{
        @NonNull
        private String name;
        List<EducationClass> classesDay = new ArrayList<>();

        public void sort(){
            classesDay = classesDay.stream().sorted().toList();
        }

        Integer[][] getRepeatClass() {
            Integer[][] result = {{0, 0}, {0, 0}, {0, 0}, {0, 0}, {0, 0}, {0, 0}, {0, 0}};
            for (EducationClass educationClass : classesDay) {
                if (educationClass.getStatusRotation() != EducationClass.RotationWeek.Continuously) {
                    int i = switch (educationClass.getStatusRotation()) {
                        case Numerator -> 0;
                        case Denominator -> 1;
                        default -> -1;
                    };
                    result[educationClass.getNumber() - 1][i] = result[educationClass.getNumber() - 1][i] + 1;
                }
            }
            return result;
        }

    }

    public void addClass(Integer day, Integer number, EducationClass.RotationWeek rotationWeek,
                         String studySubject, EducationClass.TypeSubject typeSubject, String nameEducator,
                         String studyAuditorium, EducationClass.StudyCorps studyCorps) {
        EducationClass educationClass = new EducationClass(number, rotationWeek, studySubject, nameEducator, typeSubject, studyAuditorium, studyCorps);
        week[day - 1].classesDay.add(educationClass);
    }
}
