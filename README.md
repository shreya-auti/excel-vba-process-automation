**TL;DR — What This Repository Demonstrates**

* Built an end-to-end Excel + VBA automation system replacing manual inventory tracking and document generation
* Designed defensive, modular VBA code resilient to data structure changes
* Cleaned and analyzed real-world messy datasets before automation
* Produced business-ready outputs (alerts, invoices, summaries)

Core Skills:
Excel VBA • Process Automation • Data Cleaning • Pivot Tables • Business Reporting • Defensive Programming

Flagship focus: VBA Business Process Automation
Supporting work: Data cleaning & visualization case studies

**Excel VBA Process Automation Suite**

A high-performance automation engine designed to eliminate manual data entry, ensure inventory integrity, and scale document generation. This repository demonstrates end-to-end thinking: from raw, messy data → structured analysis → automated business outputs.

**Flagship Project: VBA-Driven Business Process Automation**

Overview
This automation engine replaces error-prone manual workflows with a modular, defensive VBA system capable of handling changing data structures and real-world inconsistencies.

Key Features

* Dynamic Inventory Management: Real-time stock tracking with automated reorder alerts using search algorithms that adapt to column shifts.
  ![Inventory Alert](screenshots/invoice_template.jpg)

* Intelligent Data Sanitization: Multi-step cleaning pipeline that handles whitespace, duplicates, and inconsistencies via optimized VBA arrays.
  ![Before and After Cleaning](screenshots/data_cleaning_logic.jpg)
  
* Batch Invoice Generation: Modular engine that groups transaction data into professional PDF invoices with automated file naming and directory management.
  ![Generated Invoice Sample](screenshots/generated_invoice.jpg)
  
Technical Implementation

* Defensive Programming: Implementation of robust Error Handling (`On Error GoTo`) to ensure system stability.
* Modular Architecture: Separation of concerns using dedicated subroutines for data, logic, and UI.
* Dynamic Range Detection: Use of `CurrentRegion` and `.Find` methods to ensure the code remains "bulletproof" even if the dataset structure changes.

Sample Data: Dummy sales data used for demonstrating invoice generation and reporting.
[sample-data-sales.xlsx](sample-data/sample-data-sales.xlsx)
> Note: The included sample dataset is a representative dummy dataset created for demonstration purposes.  
> Screenshot visuals were generated from an earlier dataset with similar structure and cleaning logic.

Documentation: Detailed explanation of the project workflow, automation logic, and design decisions
[Project_Overview_Excel_VBA_Automation](docs/Project_Overview_Excel_VBA_Automation.docx)

**Supporting Module 1: Equipment Data Cleaning & Aggregation**

Purpose
Before automation and reporting, data must be trustworthy.
This module demonstrates rigorous data preparation techniques applied to operational equipment records.

What Was Done

* Removal of empty rows and duplicate records
* Text normalization and spelling correction
* Reconstruction of broken department fields using Flash Fill
* Validation using AutoSum checks (SUM, AVERAGE, MIN, MAX, COUNT)
* Multi-level pivot tables for departmental and equipment-class analysis

Outcome

Produced clean, analysis-ready datasets and structured summaries suitable for downstream automation and reporting

See: data-analysis/

**Supporting Module 2: Airbnb NYC Data Visualization Case Study**

Overview
Analyzed the 2019 Airbnb New York City dataset (~50,000 listings) to explore pricing behavior, room-type distribution, and geographic concentration.

Key Insights

* Manhattan and Brooklyn dominate Airbnb activity
* Entire homes/apartments are the most prevalent listing type
* Significant price variation across boroughs
* Lower-priced listings tend to receive more reviews, indicating higher occupancy dynamics

Tools Used

* Microsoft Excel
* Pivot Tables
* Bar, Pie, Box, and Scatter plots
* Descriptive analysis for business interpretation

See: data-visualization/

Repository Structure
src/           : Source code (.bas files) for technical review
screenshots/   : Visual proof of dashboards and generated outputs
data-analysis/ : Equipment data cleaning and pivots
docs/          : Project documentation (PDF)
sample-data/   : Normalized datasets used for stress-testing
README.md       

About the Developer

I am an Engineering Student focused on the intersection of data analytics, systems thinking, and practical business automation. 

While many focus on building complex software, I specialize in high-impact efficiency: taking messy, manual workflows and transforming them into scalable, bulletproof systems using the tools businesses use every day. 

Core Competencies:

* Data Preparation & Modeling: Turning raw, inconsistent data into reliable inputs
* Process Automation: Eliminating repetitive manual work through VBA
* Defensive Engineering: Building systems that fail gracefully and adapt to change
* Business-Oriented Analysis: Translating data into operational insigh

Let's connect: https://www.linkedin.com/in/shreya-auti-18a17537a

