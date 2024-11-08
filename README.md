# CS-Transport-Order
The Transport Order Automation Tool simplifies transport scheduling with a Python/Tkinter GUI, automating data entry into an Excel sheet by shift and route. 

## ### Project Overview

This Python and Tkinter application provides an easy-to-use interface for managing transport orders by filling in a predefined Excel table with passenger names and corresponding routes. The tool minimizes manual input by 89%, reducing human error and making the process repeatable and efficient. This project includes features to handle multiple shifts, store inputted data, clear past entries, and export a printable version.

Features
Version 1 (Initial Features)
Excel Table Integration: Automatically fills out passenger names and routes into designated sections: Morning Pickup, Night Pickup, Mid Shift Drop, and Night Drop.
GUI for Data Entry: User-friendly interface to add passenger names and their respective routes with navigation buttons to move between shift sections.
Route Collation: If multiple passengers have the same route, their names are added to the same row.

Version 2 (Enhancements)
Passenger Count & Labels: Automatically counts passengers per route (in column 6) and labels each route as "Budget Car" (in column 7).
Shift-Based Time Stamps: Automatically adds a predefined time in column 8 based on the shift type (e.g., Morning Pickup = 5:30 AM).

Version 3 (Error Handling and Clear Data Functionality)
Clear Data Button: Clears data from specified rows without affecting other data. The function also skips rows as specified.
Error Resolution: Fixed issues related to file permissions and read-only attributes when clearing data.

Version 4 (Date and Dropdown Options)
Date Integration: Fetches and displays the current date (using network connection) on the Excel sheet.
Dropdown Selection: A dropdown menu in the GUI allows the user to select one of three names, which is printed to the Excel sheet in the appropriate cell.

Version 5 (Automated Route Assignment)
Automated Route Lookup: Instead of manual route entry, the program retrieves the route information from an external Excel file ("Route.xlsx") based on the passenger name.

Version 6 (Data Synchronization & Opening Files)
Automatic Replication: Passenger entries in the "Night Shift Pickup" section are duplicated in the "Night Shift Drop" section.
File Auto-Open: After generating, the Excel file opens automatically.

Version 7 (GUI Aesthetic Improvements)
Enhanced Interface: GUI updated with a logo, orange gradient, and polished design.

![image](https://github.com/user-attachments/assets/b1412c42-aaa3-46ed-882d-d9f0c4678db1)

Future Improvements

Extend compatibility to additional Excel formats.

Add user management features to support multiple operators.

![image](https://github.com/user-attachments/assets/92f849c3-a87e-4231-9234-955356a439fc)


### Steps to follow

1. When you first open the desktop app, click on "Clear Data" to remove any previous data entries.
![image](https://github.com/user-attachments/assets/dcb0d283-1328-42b0-ad94-33326951ca19)





