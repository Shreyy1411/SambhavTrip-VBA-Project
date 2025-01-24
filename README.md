# SambhavTrip-VBA-Project
Developed a comprehensive tourism management application, “Sambhav Trip,” combining Microsoft Access, Excel, and VBA. Designed a database backend, an Excel-based user-friendly front-end, and VBA middleware to manage bookings, analyze revenue, and streamline data input.

# Sambhav Trip: Tourism Management Application

## Overview
Sambhav Trip is a tourism management application designed to streamline customer bookings, analyze revenue insights, and simplify the addition of new bookings. This project integrates Microsoft Access for backend database management, Microsoft Excel for front-end visualization, and VBA for middleware automation.

## Features
1. **Access Database**:
   - Stores customer, tour package, and booking data.
   - Includes relationships and queries for efficient data management.

2. **Excel Front-End**:
   - **BookingDetails Sheet**:
     - Displays confirmed and pending bookings with conditional formatting.
   - **RevenueInsights Sheet**:
     - Provides a pivot table summarizing revenue and booking counts by tour package.
   - **BookingInputForm Sheet**:
     - Allows users to add new bookings via a form.

3. **VBA Middleware**:
   - Automates data retrieval and insertion between Access and Excel.
   - Ensures data validation and user-friendly interactions.

## Database Design
- **Tables**:
  - Customers: Stores customer details.
  - TourPackages: Contains information about tour packages.
  - Bookings: Tracks bookings, including customer and package relationships.
- **Relationships**:
  - Customers and Bookings: Linked via `CustomerID`.
  - TourPackages and Bookings: Linked via `PackageID`.
- **Queries**:
  - `ListofActiveBookings`: Shows confirmed and pending bookings.
  - `RevenueByTourPackage`: Summarizes revenue and bookings by package.
  - `BookingsByCustomer`: Lists all bookings for a specific customer.

## Excel Workbook
- **BookingDetails**:
  - Retrieves data from the `ListofActiveBookings` query.
  - Conditional formatting: Green for confirmed, yellow for pending bookings.
- **RevenueInsights**:
  - Displays revenue summaries using a pivot table.
- **BookingInputForm**:
  - Simplifies adding bookings to the database with dropdowns and validation.

## VBA Subroutines
1. **LoadBookingDetails**:
   - Fetches and displays active bookings in the `BookingDetails` sheet.
   - Clears old data before loading new records.
2. **LoadRevenueInsights**:
   - Retrieves data from the `RevenueByTourPackage` query.
   - Updates the pivot table in the `RevenueInsights` sheet.
3. **SubmitBooking**:
   - Validates and inserts new bookings into the database.
   - Ensures error-free SQL queries and clears the input form after submission.

## Getting Started
1. Clone the repository or download the project files.
2. Open the Access database file and ensure it's accessible via your system.
3. Open the Excel workbook, enable macros, and interact with the sheets:
   - Use the `BookingDetails` sheet to view bookings.
   - Explore the `RevenueInsights` sheet for package-wise insights.
   - Use the `BookingInputForm` sheet to add new bookings.

## Dependencies
- Microsoft Office (Excel and Access).
- Macros must be enabled in Excel for VBA functionality.
