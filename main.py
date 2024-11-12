import serial
import time
import pandas as pd

# Set up serial communication
ser = serial.Serial('COM8', 9600)  # Replace 'COM5' with your Arduino's serial port
time.sleep(2)  # Wait for the connection to initialize

def enroll_fingerprint(fingerprint_id):
    # Send ENROLL command to Arduino
    ser.write(b"ENROLL\n")
    time.sleep(1)

    # Wait for Arduino response to be ready for ID
    while True:
        if ser.in_waiting > 0:
            response = ser.readline().decode(errors="ignore").strip()
            print("Arduino:", response)  # For debugging to see all messages
            if "Ready to enroll a fingerprint!" in response:
                break  # Proceed when Arduino is ready for the ID

    # Send the fingerprint ID after Arduino is ready
    ser.write(f"{fingerprint_id}\n".encode())
    time.sleep(1)

    # Wait for enrollment success or failure status from Arduino
    while True:
        if ser.in_waiting > 0:
            response = ser.readline().decode(errors="ignore").strip()
            print("Arduino:", response)  # For debugging to see all messages
            if response == "ENROLL_SUCCESS":
                print("Enrollment successful!")
                return True
            elif response == "ENROLL_FAILED":
                print("Enrollment failed.")
                return False

    print("Unknown response:", response)
    return False

def collect_additional_info():
    # Collect user details to save with fingerprint data
    details = {}
    details["ID"] = input("Enter Fingerprint ID: ")
    details["Name"] = input("Enter Name: ")
    details["DOB"] = input("Enter Date of Birth: ")
    details["Father's Name"] = input("Enter Father's Name: ")
    details["DL Number"] = input("Enter DL Number: ")
    details["DL Expiry"] = input("Enter DL Expiry Date: ")
    details["Aadhar Number"] = input("Enter Aadhar Number: ")
    return details

def save_to_excel(data, filename="C:/Users/prave/Desktop/praveen/Rnaipl/Alcohol/alcohol.xlsx"):
    # Save the details to an Excel file
    df = pd.DataFrame([data])
    try:
        existing_df = pd.read_excel(filename)
        df = pd.concat([existing_df, df], ignore_index=True)
    except FileNotFoundError:
        pass
    df.to_excel(filename, index=False)
    print(f"Data saved to {filename}")

if __name__ == "__main__":
    while True:
        fingerprint_id = int(input("Enter the ID for the fingerprint: "))
        if enroll_fingerprint(fingerprint_id):
            # Only collect details and save if enrollment was successful
            additional_info = collect_additional_info()
            save_to_excel(additional_info)
        else:
            print("Fingerprint enrollment was not successful. Try again.")

        # Ask if the user wants to enroll another fingerprint or stop
        user_input = input("Do you want to enroll another fingerprint? (y/n): ").lower()
        if user_input == "n" or user_input == "stop":
            print("Exiting enrollment process.")
            break  # Exit the loop

    # Close the serial connection after completing the enrollment process
    ser.close()
    print("Serial connection closed.")
