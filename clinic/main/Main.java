package clinic.main;
import clinic.ClinicManagementSystem;
import java.util.Scanner;

public class Main {

    public static void main(String[] args) {
        ClinicManagementSystem cms = new ClinicManagementSystem();
        Scanner scanner = new Scanner(System.in);
        boolean running = true;

        System.out.println("\nWelcome to the Clinic Management System");

        while (running) {
            displayMenu();
            System.out.print("Enter choice: ");
            String choiceStr = scanner.nextLine();

            try {
                int choice = Integer.parseInt(choiceStr);
                switch (choice) {
                    case 1:
                        cms.checkInPatient(scanner);
                        break;
                    case 2:
                        cms.treatNextPatient(scanner);
                        break;
                    case 3:
                        cms.viewAllPatientHistory(scanner);
                        break;
                    case 4:
                        cms.viewDoctorStatus(scanner);
                        break;
                    case 5:
                        runSetupProcedures(cms, scanner);
                        break;
                    case 6:
                        cms.exitAndSave();
                        running = false;
                        break;
                    default:
                        System.out.println("Invalid choice. Please enter a number between 1 and 6.");
                }
            } catch (NumberFormatException e) {
                System.out.println("Invalid input. Please enter a number.");
            }
        }
        scanner.close();
    }

    private static void displayMenu() {
        System.out.println("\n======== Main Menu ===========");
        System.out.println("[1] Set Appointment");
        System.out.println("[2] Treat Next Patient");
        System.out.println("[3] View Patient History");
        System.out.println("[4] Doctor List");
        System.out.println("[5] Setup Procedures");
        System.out.println("[6] Exit Program");
        System.out.println("==============================");
    }

    private static void runSetupProcedures(ClinicManagementSystem cms, Scanner scanner) {
        boolean inSetup = true;
        while (inSetup) {
            System.out.println("\n==== Setup Procedures ===");
            System.out.println("[1] Register New Doctor");
            System.out.println("[2] Remove Existing Doctor");
            System.out.println("[3] Add New Room");
            System.out.println("[4] Remove Existing Room");
            System.out.println("[5] Back to Main Menu");
            System.out.println("=========================");
            System.out.print("Enter choice: ");

            try {
                int choice = Integer.parseInt(scanner.nextLine());
                switch (choice) {
                    case 1:
                        cms.registerNewDoctor(scanner);
                        break;
                    case 2:
                        cms.removeDoctor(scanner);
                        break;
                    case 3:
                        cms.addTreatmentRoom(scanner);
                        break;
                    case 4:
                        cms.removeRoom(scanner);
                        break;
                    case 5:
                        inSetup = false;
                        break;
                    default:
                        System.out.println("Invalid choice.");
                }
            } catch (NumberFormatException e) {
                System.out.println("Invalid input.");
            }
        }
    }
}
