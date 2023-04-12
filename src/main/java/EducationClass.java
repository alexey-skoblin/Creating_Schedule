import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import java.util.Objects;

@Getter
@Setter
@Slf4j
@AllArgsConstructor
public class EducationClass implements Comparable<EducationClass> {
    private Integer number;
    private RotationWeek statusRotation;
    private String studySubject;
    private String nameEducator;
    private TypeSubject typeSubject;
    private String studyAuditorium;
    private StudyCorps studyCorp;

    @Override
    public int compareTo(EducationClass educationClass) {
        if (this.number < educationClass.getNumber())
            return -1;
        else if (this.number > educationClass.getNumber())
            return 1;
        else {
            if (this.statusRotation == RotationWeek.Continuously && educationClass.getStatusRotation() == RotationWeek.Continuously)
                return 0;
            else if (this.statusRotation == RotationWeek.Numerator && educationClass.getStatusRotation() == RotationWeek.Numerator)
                return 0;
            else if (this.statusRotation == RotationWeek.Denominator && educationClass.getStatusRotation() == RotationWeek.Denominator)
                return 0;
            else if (this.statusRotation == RotationWeek.Numerator && educationClass.getStatusRotation() == RotationWeek.Continuously)
                return 1;
            else if (this.statusRotation == RotationWeek.Denominator && educationClass.getStatusRotation() == RotationWeek.Numerator)
                return 1;
            else if (this.statusRotation == RotationWeek.Denominator && educationClass.getStatusRotation() == RotationWeek.Continuously)
                return 1;
            else return -1;
        }
    }

    enum RotationWeek {
        Continuously,
        Numerator,
        Denominator
    }

    static RotationWeek convertStringToRotationWeek(String s) {
        return switch (s) {
            case "-" -> RotationWeek.Continuously;
            case "Ч" -> RotationWeek.Numerator;
            case "З" -> RotationWeek.Denominator;
            default -> null;
        };
    }

    enum TypeSubject {
        Lecture,
        Practice,
        Laboratory
    }

    static TypeSubject convertStringToTypeSubject(String s) {
        return switch (s) {
            case "(Лек)" -> TypeSubject.Lecture;
            case "(Прак)" -> TypeSubject.Practice;
            case "(Лаб)" -> TypeSubject.Laboratory;
            default -> null;
        };
    }

    enum StudyCorps {
        Corp1,
        Corp4,
        Corp6,
        Library
    }

    static StudyCorps convertStringToStudyCorp(String s) {
        return switch (s) {
            case "1к" -> StudyCorps.Corp1;
            case "4к" -> StudyCorps.Corp4;
            case "6к" -> StudyCorps.Corp6;
            case "НБ" -> StudyCorps.Library;
            default -> null;
        };
    }

    static String[] ScheduleCalls = {
            "8:20-9.50",
            "10:00-11:30",
            "12:00-13:30",
            "13:40-15:10",
            "15:20-16:50",
            "17:10-18:40",
            "18:50-20:20"
    };
}
