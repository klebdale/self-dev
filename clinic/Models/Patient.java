package clinic.Models;

public class Patient implements Comparable<Patient> {
    private final String patientID;
    private final String name;
    private final String concern;
    private final int priorityLevel;
    private String roomID = "N/A";
    private String doctorID = "N/A";

    public Patient(String patientID, String name, String concern, int priorityLevel) {
        this.patientID = patientID;
        this.name = name;
        this.concern = concern;
        this.priorityLevel = priorityLevel;
    }

    public String getPatientID() {
        return patientID;
    }

    public String getName() {
        return name;
    }

    public int getPriorityLevel() {
        return priorityLevel;
    }

    public String getConcern() {
        return concern;
    }

    public void setDoctorID(String doctorID) {
        this.doctorID = doctorID;
    }

    public void  setRoomID(String roomID) {
        this.roomID = roomID;
    }

    public String getDoctorID() {
        return doctorID;
    }

    @Override
    public int compareTo(Patient other) {
        return Integer.compare(this.priorityLevel, other.priorityLevel);    }

    @Override
    public String toString() {
        return name + " (P: " + priorityLevel + ", Concern: " + concern + ")";
    }

}
