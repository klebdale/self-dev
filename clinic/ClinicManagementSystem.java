package clinic;

import clinic.Models.*;
import java.util.*;
import java.io.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.concurrent.ThreadLocalRandom;

public class ClinicManagementSystem {

    // ------------------- FILE LOCATIONS -------------------
    private final String DATA_FILE = "clinic_data.csv";
    private final DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    // Data Structures here
    private final PriorityQueue<Patient> waitingQueue = new PriorityQueue<>();
    private final HashMap<String, List<ClinicHistoryRecord>> patientHistoryMap = new HashMap<>();
    private final HashMap<String, String> patientNames = new HashMap<>();
    private final HashMap<String, Doctor> doctorMap = new HashMap<>();
    private final HashMap<String, Room> roomMap = new HashMap<>();

    // ------------------- CONSTRUCTOR -------------------
    public ClinicManagementSystem() {
        loadAllData();
    }

    private String generateShortID(String prefix) {
        String shortId;
        int attempts = 0;
        do {
            long randomBytes = ThreadLocalRandom.current().nextLong(0, 0x100000000L);
            String hexPart = String.format("%08x", randomBytes);
            shortId = prefix + hexPart;
            attempts++;

            if (prefix.equals("P") && !patientNames.containsKey(shortId)) return shortId;
            if (prefix.equals("D") && !doctorMap.containsKey(shortId)) return shortId;
            if (prefix.equals("R") && !roomMap.containsKey(shortId)) return shortId;
        } while (attempts < 10);
        throw new RuntimeException("Failed to generate unique short ID.");
    }

    public void checkInPatient(Scanner scanner) {
        System.out.println("\n--- Patient Check-In ---");

        List<Doctor> availableDoctors = getAvailableDoctors();
        if (availableDoctors.isEmpty()) {
            System.out.println("\n!! No doctors are currently available (In Clinic). Cannot check in patient !!");
            return;
        }

        System.out.print("Enter Patient's Name: ");
        String name = scanner.nextLine();
        System.out.print("Enter Patient's Concern: ");
        String concern = scanner.nextLine();

        int priorityLevel = 0;
        while (priorityLevel < 1 || priorityLevel > 5) {
            try {
                System.out.print("Enter Priority Level ( 1 = Urgent, 5 = Least Urgent): ");
                priorityLevel = Integer.parseInt(scanner.nextLine());
            } catch (NumberFormatException e) {
                System.out.println("\n!! Invalid input. Please enter a number !!");
            }
        }

        System.out.println("\n----- Assign a Doctor -----");
        for (int i = 0; i < availableDoctors.size(); i++) {
            System.out.println((i + 1) + ". " + availableDoctors.get(i).getName());
        }

        Doctor assignedDoctor = null;
        while (assignedDoctor == null) {
            System.out.print("Select Doctor\n(Enter Number): ");
            try {
                int choice = Integer.parseInt(scanner.nextLine());
                if (choice > 0 && choice <= availableDoctors.size()) {
                    assignedDoctor = availableDoctors.get(choice - 1);
                } else {
                    System.out.println("\n!! Invalid selection !!");
                }
            } catch (NumberFormatException e) {
                System.out.println("\n!! Invalid input !!");
            }
        }

        String patientID = generateShortID("P");
        Patient newPatient = new Patient(patientID, name, concern, priorityLevel);
        newPatient.setDoctorID(assignedDoctor.getDoctorId());

        waitingQueue.add(newPatient);
        patientNames.put(patientID, name);

        System.out.println("\n=============== SUCCESS ===================");
        System.out.println("Patient " + name + " checked in successfully!");
        System.out.println("Assigned to: " + assignedDoctor.getName());
        System.out.println("=============================================");
    }

    public void treatNextPatient(Scanner scanner) {
        if (waitingQueue.isEmpty()) {
            System.out.println("\n---------- ERROR ----------");
            System.out.println("!! NO PATIENTS IN QUEUE !!");
            System.out.println("---------------------------");
            return;
        }

        Patient currentPatient = waitingQueue.poll();
        System.out.println("\n--- Treating Patient: " + currentPatient.getName() + " (P" + currentPatient.getPriorityLevel() + ") ---");

        String assignedDocID = currentPatient.getDoctorID();
        Doctor assignedDoctor = doctorMap.get(assignedDocID);

        if (!assignedDoctor.isInClinic()) {
            System.out.println("\n------------- ATTENTION --------------");
            System.out.println("Assigned Doctor (" + assignedDoctor.getName() + ") is currently UNAVAILABLE/BUSY.");
            System.out.println("1. Assign a new available doctor");
            System.out.println("2. Wait for " + assignedDoctor.getName() + " (Return Patient to Queue)");
            System.out.println("--------------------------------------");
            System.out.print("Enter choice: ");

            int choice = -1;
            try {
                choice = Integer.parseInt(scanner.nextLine());
            } catch (NumberFormatException e) {
                System.out.println("Invalid input. Defaulting to wait.");
            }

            if (choice == 1) {
                List<Doctor> availableDoctors = getAvailableDoctors();

                if (availableDoctors.isEmpty()) {
                    System.out.println("\n------------------- ATTENTION ---------------------");
                    System.out.println("Sorry, there are NO other doctors available either.");
                    System.out.println("Returning patient to queue.");
                    System.out.println("---------------------------------------------------");
                    waitingQueue.add(currentPatient);
                    return;
                }

                System.out.println("\n--- Available Doctors ---");
                for (int i = 0; i < availableDoctors.size(); i++) {
                    System.out.println((i + 1) + ". " + availableDoctors.get(i).getName());
                }

                Doctor newDoctor = null;
                while (newDoctor == null) {
                    System.out.print("Select New Doctor (Enter Number): ");
                    try {
                        int docChoice = Integer.parseInt(scanner.nextLine());
                        if (docChoice > 0 && docChoice <= availableDoctors.size()) {
                            newDoctor = availableDoctors.get(docChoice - 1);
                        } else {
                            System.out.println("Invalid selection.");
                        }
                    } catch (NumberFormatException e) {
                        System.out.println("Invalid input.");
                    }
                }

                currentPatient.setDoctorID(newDoctor.getDoctorId());
                assignedDoctor = newDoctor;
                System.out.println("\n-------------- Success --------------");
                System.out.println("Doctor successfully changed to " + assignedDoctor.getName());
                System.out.println("-------------------------------------");

            } else {
                System.out.println("\n-------------------- Returning --------------------");
                System.out.println("Returning patient to queue to wait for " + assignedDoctor.getName() + ".");
                System.out.println("---------------------------------------------------");
                waitingQueue.add(currentPatient);
                return;
            }
        }

        Room availableRoom = findAvailableRoom();
        if (availableRoom == null) {
            System.out.println("\n--------------- ERROR ----------------");
            System.out.println("No rooms available. Returned to queue.");
            System.out.println("--------------------------------------");
            waitingQueue.add(currentPatient);
            return;
        }

        assignedDoctor.setInClinic(false);
        availableRoom.setOccupied(true);

        System.out.println("Assigned Doctor: " + assignedDoctor.getName());
        System.out.println("Assigned Room: " + availableRoom.getType());

        System.out.print("Enter Diagnosis: ");
        String diagnosis = scanner.nextLine();
        System.out.print("Enter Treatment Summary: ");
        String treatment = scanner.nextLine();

        LocalDateTime now = LocalDateTime.now();
        ClinicHistoryRecord record = new ClinicHistoryRecord(now, assignedDoctor.getName(), diagnosis, treatment);

        patientHistoryMap.computeIfAbsent(currentPatient.getPatientID(), k -> new ArrayList<>()).add(record);

        assignedDoctor.setInClinic(true);
        availableRoom.setOccupied(false);

        System.out.println("\n----------- Success --------------");
        System.out.println("Treatment Complete. History saved.");
        System.out.println("----------------------------------");
    }

    public void viewAllPatientHistory(Scanner scanner) {
        System.out.println("\n--- View Patient History ---");

        if (patientNames.isEmpty()) {
            System.out.println("\n---------- Attention -------------");
            System.out.println("No patients found in the system.");
            System.out.println("----------------------------------");
            return;
        }

        List<String> patientIDs = new ArrayList<>(patientNames.keySet());

        System.out.println("Select a patient to view records:");
        for (int i = 0; i < patientIDs.size(); i++) {
            String id = patientIDs.get(i);
            String name = patientNames.get(id);
            System.out.println((i + 1) + ". " + name + " (ID: " + id + ")");
        }

        System.out.print("Enter number (or 0 to cancel): ");
        try {
            int choice = Integer.parseInt(scanner.nextLine());
            if (choice == 0) return;

            if (choice > 0 && choice <= patientIDs.size()) {
                String selectedID = patientIDs.get(choice - 1);
                String selectedName = patientNames.get(selectedID);

                List<ClinicHistoryRecord> records = patientHistoryMap.get(selectedID);

                System.out.println("\n--- Medical History: " + selectedName + " ---");
                if (records == null || records.isEmpty()) {
                    System.out.println("No completed treatment records found for this patient yet.");
                } else {
                    for (ClinicHistoryRecord record : records) {
                        System.out.println(record.toString());
                        System.out.println("-------------------------");
                    }
                }
            } else {
                System.out.println("Invalid selection.");
            }
        } catch (NumberFormatException e) {
            System.out.println("Invalid input.");
        }
    }

    public void viewDoctorStatus(Scanner scanner) {
        if (doctorMap.isEmpty()) {
            return;
        }

        System.out.println("\n--- Doctor List ---");
        List<Doctor> doctorList = new ArrayList<>(doctorMap.values());

        for (int i = 0; i < doctorList.size(); i++) {
            System.out.println((i + 1) + ". " + doctorList.get(i));
        }

        System.out.print("\nSelect doctor to change status (or 0 to cancel): ");
        try {
            int choice = Integer.parseInt(scanner.nextLine());
            if (choice > 0 && choice <= doctorList.size()) {
                Doctor selectedDoc = doctorList.get(choice - 1);

                System.out.println("Current Status: " + (selectedDoc.isInClinic() ? "In Clinic" : "Busy"));
                System.out.println("1. Set to In Clinic (Available)");
                System.out.println("2. Set to Busy (Unavailable)");
                System.out.print("Choose status: ");

                int statusChoice = Integer.parseInt(scanner.nextLine());
                if (statusChoice == 1) {
                    selectedDoc.setInClinic(true);
                    System.out.println(selectedDoc.getName() + " is now Available.");
                } else if (statusChoice == 2) {
                    selectedDoc.setInClinic(false);
                    System.out.println(selectedDoc.getName() + " is now Busy.");
                } else {
                    System.out.println("Invalid status choice.");
                }
            }
        } catch (NumberFormatException e) {
            System.out.println("Invalid input.");
        }
    }

    public void registerNewDoctor(Scanner scanner) {
        System.out.println("\n--- Register New Doctor ---");
        System.out.print("Enter Doctor's Name: ");
        String name = scanner.nextLine();

        String doctorID = generateShortID("D");
        Doctor newDoctor = new Doctor(doctorID, name, true);
        doctorMap.put(doctorID, newDoctor);

        System.out.println("------------- SUCCESS --------------");
        System.out.println("Doctor " + name + " registered.");
        System.out.println("------------------------------------");
    }

    public void removeDoctor(Scanner scanner) {
        System.out.println("\n--- Remove Doctor ---");
        if (doctorMap.isEmpty()) return;

        List<Doctor> doctorList = new ArrayList<>(doctorMap.values());
        for (int i = 0; i < doctorList.size(); i++) {
            System.out.println((i + 1) + ". " + doctorList.get(i));
        }

        System.out.print("\nEnter number to remove (or 0 to cancel): ");
        try {
            int choice = Integer.parseInt(scanner.nextLine());
            if (choice > 0 && choice <= doctorList.size()) {
                Doctor toRemove = doctorList.get(choice - 1);
                doctorMap.remove(toRemove.getDoctorId());
                System.out.println("Removed " + toRemove.getName());
            }
        } catch (NumberFormatException e) {
            System.out.println("Invalid input.");
        }
    }

    public void addTreatmentRoom(Scanner scanner) {
        System.out.println("\n--- Add New Room ---");
        System.out.print("Enter Room Name: ");
        String type = scanner.nextLine();

        String roomID = generateShortID("R");
        Room newRoom = new Room(roomID, type, false);
        roomMap.put(roomID, newRoom);

        System.out.println("\n--------------- SUCCESS ---------------");
        System.out.println("SUCCESS: Room " + roomID + " added.");
        System.out.println("---------------------------------------");
    }

    public void removeRoom(Scanner scanner) {
        System.out.println("\n--- Remove Room ---");
        if (roomMap.isEmpty()) return;

        List<Room> roomList = new ArrayList<>(roomMap.values());
        for (int i = 0; i < roomList.size(); i++) {
            System.out.println((i + 1) + ". " + roomList.get(i));
        }

        System.out.print("\nEnter number to remove (or 0 to cancel): ");
        try {
            int choice = Integer.parseInt(scanner.nextLine());
            if (choice > 0 && choice <= roomList.size()) {
                Room toRemove = roomList.get(choice - 1);
                roomMap.remove(toRemove.getRoomID());
                System.out.println("Removed Room " + toRemove.getType());
            }
        } catch (NumberFormatException e) {
            System.out.println("Invalid input.");
        }
    }

    private List<Doctor> getAvailableDoctors() {
        List<Doctor> available = new ArrayList<>();
        for (Doctor d : doctorMap.values()) {
            if (d.isInClinic()) {
                available.add(d);
            }
        }
        return available;
    }

    private Room findAvailableRoom() {
        for (Room r : roomMap.values()) {
            if (!r.isOccupied()) return r;
        }
        return null;
    }

    // ------------------- FILE HANDLING -------------------
    public void exitAndSave() {
        saveAllData();
        System.out.println("System shutting down. Goodbye!");
    }

    private void saveAllData() {
        try (PrintWriter pw = new PrintWriter(new FileWriter(DATA_FILE))) {
            // PATIENTS SECTION
            pw.println("[PATIENTS]");
            pw.println("PatientID,Name");
            for (String id : patientNames.keySet()) {
                pw.println(id + "," + patientNames.get(id));
            }
            pw.println();

            // DOCTORS SECTION
            pw.println("[DOCTORS]");
            pw.println("DoctorID,Name,InClinic");
            for (Doctor d : doctorMap.values()) {
                pw.println(d.getDoctorId() + "," + d.getName() + "," + d.isInClinic());
            }
            pw.println();

            // ROOMS SECTION
            pw.println("[ROOMS]");
            pw.println("RoomID,Type,IsOccupied");
            for (Room r : roomMap.values()) {
                pw.println(r.getRoomID() + "," + r.getType() + "," + r.isOccupied());
            }
            pw.println();

            // WAITING QUEUE SECTION
            pw.println("[QUEUE]");
            pw.println("PatientID,Name,Concern,Priority,DoctorID");
            for (Patient p : waitingQueue) {
                pw.println(p.getPatientID() + "," + 
                          p.getName() + "," + 
                          p.getConcern() + "," + 
                          p.getPriorityLevel() + "," + 
                          p.getDoctorID());
            }
            pw.println();

            // HISTORY SECTION
            pw.println("[HISTORY]");
            pw.println("PatientID,DateTime,Doctor,Diagnosis,Treatment");
            for (String id : patientHistoryMap.keySet()) {
                for (ClinicHistoryRecord r : patientHistoryMap.get(id)) {
                    pw.println(
                            id + "," +
                            r.getDateTime().format(formatter) + "," +
                            r.getDoctorName() + "," +
                            r.getDiagnosis().replace(",", ";") + "," +
                            r.getTreatment().replace(",", ";")
                    );
                }
            }

        } catch (IOException e) {
            System.err.println("Error saving data: " + e.getMessage());
        }
    }

    private void loadAllData() {
        File file = new File(DATA_FILE);
        if (!file.exists()) {
            return;
        }

        try (BufferedReader br = new BufferedReader(new FileReader(file))) {
            String line;
            String currentSection = "";

            while ((line = br.readLine()) != null) {
                line = line.trim();
                
                if (line.startsWith("[") && line.endsWith("]")) {
                    currentSection = line.substring(1, line.length() - 1);
                    br.readLine(); // Skip column headers
                    continue;
                }

                if (line.isEmpty()) {
                    continue;
                }

                String[] parts = line.split(",", -1);

                switch (currentSection) {
                    case "PATIENTS":
                        if (parts.length >= 2) {
                            patientNames.put(parts[0], parts[1]);
                        }
                        break;

                    case "DOCTORS":
                        if (parts.length >= 3) {
                            String id = parts[0];
                            String name = parts[1];
                            boolean inClinic = Boolean.parseBoolean(parts[2]);
                            doctorMap.put(id, new Doctor(id, name, inClinic));
                        }
                        break;

                    case "ROOMS":
                        if (parts.length >= 3) {
                            String id = parts[0];
                            String type = parts[1];
                            boolean occupied = Boolean.parseBoolean(parts[2]);
                            roomMap.put(id, new Room(id, type, occupied));
                        }
                        break;

                    case "QUEUE":
                        if (parts.length >= 5) {
                            String id = parts[0];
                            String name = parts[1];
                            String concern = parts[2];
                            int priority = Integer.parseInt(parts[3]);
                            String doctorID = parts[4];

                            Patient patient = new Patient(id, name, concern, priority);
                            patient.setDoctorID(doctorID);
                            waitingQueue.add(patient);
                        }
                        break;

                    case "HISTORY":
                        if (parts.length >= 5) {
                            String id = parts[0];
                            LocalDateTime time = LocalDateTime.parse(parts[1], formatter);
                            ClinicHistoryRecord r = new ClinicHistoryRecord(time, parts[2], parts[3], parts[4]);
                            patientHistoryMap.computeIfAbsent(id, k -> new ArrayList<>()).add(r);
                        }
                        break;
                }
            }

        } catch (IOException e) {
            System.err.println("Error loading data: " + e.getMessage());
        }
    }
}
