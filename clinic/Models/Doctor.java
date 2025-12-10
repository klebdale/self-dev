package clinic.Models;

public class Doctor {

    private final String doctorId;
    private final String name;
    private boolean isInClinic;

    public Doctor(String doctorId, String name, boolean isInClinic) {
        this.doctorId = doctorId;
        this.name = name;
        this.isInClinic = isInClinic;
    }

    public String getDoctorId() {
        return doctorId;
    }

    public String getName() {
        return name;
    }

    public boolean isInClinic() {
        return isInClinic;
    }

    public void setInClinic(boolean inClinic) {
        isInClinic = inClinic;
    }

    @Override
    public String toString() {
        return name + " - ID: " + doctorId + (isInClinic ? " [In Clinic]" : " [Busy]");
    }
}
