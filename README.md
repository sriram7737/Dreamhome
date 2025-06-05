# Dreamhome Property Viewing System

## Overview

Dreamhome is a classic ASP-based web application designed to manage property viewings for a real estate agency. It facilitates functionalities such as client registration, property viewing bookings, and administrative review of scheduled viewings.

##  Features

- **Client Registration**: Allows new clients to register their details.
- **Property Viewing Booking**: Enables clients to schedule viewings for available properties.
- **Administrative Review**: Provides administrators with tools to review and manage scheduled viewings.
- **Email Notifications**: Sends confirmation emails upon successful booking.

##  Project Structure

Dreamhome/
├── AddToMail.asp # Handles adding clients to the mailing list
├── AddToMail.htm # Frontend form for mailing list subscription
├── BookViewing.asp # Processes property viewing bookings
├── BookViewing.htm # Frontend form for booking viewings
├── DHFunctions.inc # Includes reusable functions
├── DreamHome Case Study.doc # Documentation of the case study
├── DreamHome.mdb # Microsoft Access database
├── Images/ # Directory containing property images
├── ReviewViewings.asp # Admin page to review scheduled viewings
├── SetClient.asp # Sets client session variables
├── UpdateViewRecord.asp # Updates viewing records
├── ViewDetails.asp # Displays detailed property information
├── createBooking.asp # Handles creation of new bookings
├── global.asa # Global ASP configuration file
├── index.asp # Main landing page
├── index.htm # Static version of the landing page
├── login.asp # Client login page
├── logout.asp # Client logout page
├── README.md # Project documentation
├── ReadThis.txt # Additional notes and instructions



##  Installation and Setup

1. **Prerequisites**:
   - Windows OS with IIS (Internet Information Services) enabled.
   - Microsoft Access installed for handling the `.mdb` database.

2. **Setup Instructions**:
   - Clone or download the repository to your local machine.
   - Place the project folder in the IIS root directory (e.g., `C:\inetpub\wwwroot\Dreamhome`).
   - Configure IIS to recognize `.asp` files and set up the application.
   - Ensure that the `DreamHome.mdb` database is accessible and properly connected within the application files.

3. **Running the Application**:
   - Start IIS and navigate to `http://localhost/Dreamhome/index.asp` in your web browser.
   - Use the provided forms to register clients, book property viewings, and manage administrative tasks.

##  Documentation

Refer to the `DreamHome Case Study.doc` file for a detailed overview of the project's objectives, functionalities, and implementation details.

## Contributing

Contributions are welcome! If you have suggestions or improvements, please fork the repository and submit a pull request.


---

For more projects and information, visit [Sriram Rampelli's GitHub Profile](https://github.com/sriram7737)
