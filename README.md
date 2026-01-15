Excel VBA Process Automation Suite

A high-performance automation engine designed to eliminate manual data entry, ensure inventory integrity, and scale document generation. This project demonstrates modular programming, defensive coding, and dynamic data handling.

Key Features

* Dynamic Inventory Management: Real-time stock tracking with automated reorder alerts using search algorithms that adapt to column shifts.
* Intelligent Data Sanitization: Multi-step cleaning pipeline that handles whitespace, duplicates, and inconsistencies via optimized VBA arrays.
* Batch Invoice Generation: Modular engine that groups transaction data into professional PDF invoices with automated file naming and directory management.

Technical Implementation

* Defensive Programming: Implementation of robust Error Handling (`On Error GoTo`) to ensure system stability.
* Modular Architecture: Separation of concerns using dedicated subroutines for data, logic, and UI.
* Dynamic Range Detection: Use of `CurrentRegion` and `.Find` methods to ensure the code remains "bulletproof" even if the dataset structure changes.

Project Structure
/src          : Source code (.bas files) for technical review
/screenshots  : Visual proof of dashboards and generated outputs
/sample-data  : Normalized datasets used for stress-testing
