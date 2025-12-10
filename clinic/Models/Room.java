package clinic.Models;

public class Room {

    private final String roomID;
    private final String type;
    private boolean isOccupied;

    public Room(String roomID, String type, boolean isOccupied) {
        this.roomID = roomID;
        this.type = type;
        this.isOccupied = isOccupied;
    }

    public String getRoomID() {
        return roomID;
    }

    public String getType() {
        return type;
    }

    // FIX: Standard boolean naming
    public boolean isOccupied() {
        return isOccupied;
    }

    // FIX: Standard boolean naming
    public void setOccupied(boolean occupied) {
        isOccupied = occupied;
    }

    @Override
    public String toString() {
        return "Room " + roomID + " (" + type + ") - " + (isOccupied ? "OCCUPIED" : "AVAILABLE");
    }
}