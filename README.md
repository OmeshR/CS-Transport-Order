# CS-Transport-Order
The Transport Order Automation Tool simplifies transport scheduling with a Python/Tkinter GUI, automating data entry into an Excel sheet by shift and route. 

## Project Overview
This Python and Tkinter application provides an easy-to-use interface for managing transport orders by filling in a predefined Excel table with passenger names and corresponding routes. The tool minimizes manual input by 89%, reducing human error and making the process repeatable and efficient. This project includes features to handle multiple shifts, store inputted data, clear past entries, and export a printable version.

## Features
**Version 1.0** (Initial Features)
*Excel Table Integration: Automatically fills out passenger names and routes into designated sections: Morning Pickup, Night Pickup, Mid Shift Drop, and Night Drop.
GUI for Data Entry: User-friendly interface to add passenger names and their respective routes with navigation buttons to move between shift sections.
Route Collation: If multiple passengers have the same route, their names are added to the same row.*

**Version 2.0** (Enhancements)
*Passenger Count & Labels: Automatically counts passengers per route (in column 6) and labels each route as "Budget Car" (in column 7).
Shift-Based Time Stamps: Automatically adds a predefined time in column 8 based on the shift type (e.g., Morning Pickup = 5:30 AM).*

**Version 3.0** (Error Handling and Clear Data Functionality)
*Clear Data Button: Clears data from specified rows without affecting other data. The function also skips rows as specified.
Error Resolution: Fixed issues related to file permissions and read-only attributes when clearing data.*

**Version 4.0** (Date and Dropdown Options)
*Date Integration: Fetches and displays the current date (using network connection) on the Excel sheet.
Dropdown Selection: A dropdown menu in the GUI allows the user to select one of three names, which is printed to the Excel sheet in the appropriate cell.*

**Version 5.0** (Automated Route Assignment)
*Automated Route Lookup: Instead of manual route entry, the program retrieves the route information from an external Excel file ("Route.xlsx") based on the passenger name.*

**Version 6.0** (Data Synchronization & Opening Files)
*Automatic Replication: Passenger entries in the "Night Shift Pickup" section are duplicated in the "Night Shift Drop" section.
File Auto-Open: After generating, the Excel file opens automatically.*

**Version 7.0** (GUI Aesthetic Improvements)
*Enhanced Interface: GUI updated with a logo, orange gradient, and polished design.*

## Future Improvements
Extend compatibility to additional Excel formats.

Add user management features to support multiple operators.

## User Interface
![image](https://github.com/user-attachments/assets/92f849c3-a87e-4231-9234-955356a439fc)


## Steps to follow
1. When you first open the desktop app, click on "Clear Data" to remove any previous data entries.
![image](https://github.com/user-attachments/assets/dcb0d283-1328-42b0-ad94-33326951ca19)

2. Check the top title to know which shift you're working on.
![image](https://github.com/user-attachments/assets/b4308e26-f480-4c5f-97b3-671b7a6f9045)

3. The type of shift (Morning pickup/ Night shift pickup/ Mid shift drop/ Night shift drop) can be changed by using "Next Shift" or "Previous Shift".
   
![image](https://github.com/user-attachments/assets/7a76c33e-c30b-41d8-bf50-84a4f7a247e0)
![image](https://github.com/user-attachments/assets/b80ecf0b-4794-4494-94e6-b3065b083170)

5. Add the passenger name based on the shift for that particular day (make sure to enter the name properly with no spelling errors, always start a name with a capital letter and only the first name should be typed).
![image](https://github.com/user-attachments/assets/cc9bf1de-4734-4a42-afa4-b75e0d5da6fd)

6. Use the "Select Name" function to add the name of the person who's making the transport order.
![image](https://github.com/user-attachments/assets/214f9b5a-11d8-4bbe-a165-b3c7e1c8335b)

7. Then click on "Add Passenger" to include the person in that day's transport order (The added name will disappear).
![image](https://github.com/user-attachments/assets/3ebe9002-1e31-49c9-8b47-507b0b4325fa)

8. Once you're done with adding people based on their shifts, click on "Generate" to create an excel table with the transport order (This excel sheet will pop-up as soon as you hit "Generate" (For any dialogue box that opens along with the excel sheet, just press on "okay")
![image](https://github.com/user-attachments/assets/dae85f35-2a61-4ffe-8c7c-5cd03a9ae63d)
![image](https://github.com/user-attachments/assets/a51c2a42-e5b7-4855-b12a-ac577090a259)

9. Use the temporarily created excel sheet to send out the transport order to the cab company. REFER BELOW TO SEE WHAT INFORMATION WILl BE ADDED AUTOMATICALLY.
![image](https://github.com/user-attachments/assets/37e5bc13-0a47-412a-8d3f-1a9ecee0ad1c)


[!IMPORTANT]

Based on the passenger's name, the route will be automatically added in front of their name.

If there are two or more passengers with the same route, the program will add them to the same transport.

Based on the number of people in one route, the program will add either "1" or "2" or "3" to the "No. of Passengers" column.

Based on the day you're working on this, the date will be automatically added.

Based on the number of routes, the program will automatically calculate the "No. of cabs" and print the sum on the table.

The name of the person who made the transport order can be selected before clicking on "Generate" and that will be printed on the table too.

Please keep your internet connection active when using the tool as it is needd to find the correct date.







