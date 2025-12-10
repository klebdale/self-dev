package clinic.Models;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class ClinicHistoryRecord {

    private final LocalDateTime timestamp;
    private final String doctorName;
    private final String diagnosis;
    private final String treatment;

    private static final DateTimeFormatter FORMATTER =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    public ClinicHistoryRecord(LocalDateTime timestamp,
                               String doctorName,
                               String diagnosis,
                               String treatment) {
        this.timestamp = timestamp;
        this.doctorName = doctorName;
        this.diagnosis = diagnosis;
        this.treatment = treatment;
    }

    // All getter methods needed
    public LocalDateTime getDateTime() {
        return timestamp;
    }

    public String getDoctorName() {
        return doctorName;
    }

    public String getDiagnosis() {
        return diagnosis;
    }

    public String getTreatment() {
        return treatment;
    }

    @Override
    public String toString() {
        return  "Date: " + timestamp.format(FORMATTER) + "\n" +
                "Doctor: " + doctorName + "\n" +
                "Diagnosis: " + diagnosis + "\n" +
                "Treatment: " + treatment;
    }
}
