# GradeReportsGoogleClassroom
Appscript which can be attached to Google Sheets for generating Google Classroom Grade Reports.

Copy the script contents from: 

Google_Classroom_Grade_Report_Generator.gs

into an appscript script attached to a google sheet.

# Google Classroom Grade Report Generator

**Version:** 1.2.4  
**Author:** William Openshaw (OphaTapioka)  
**Last Updated:** December 05, 2025

## Overview

The **Google Classroom Grade Report Generator** is a Google Apps Script designed to automate the creation of detailed grade reports for students. It fetches data directly from Google Classroom, calculates grades based on weighted categories, and generates individual report sheets for each student. It also includes functionality to email these reports as PDFs to students and their guardians.

This tool is specifically designed for classes that use **weighted grading categories** (e.g., Homework 20%, Tests 40%, etc.).

## Features

* **Automated Data Fetching**: Retrieves courses, assignments, and grades from Google Classroom.
* **Weighted Grading Support**: Specifically handles logic for courses graded by weighted categories.
* **Individual Student Reports**: Generates a dedicated tab in the Google Sheet for every student, organizing their grades by category.
* **PDF Generation & Emailing**: Converts student sheets into PDF reports and emails them to students and guardians.
* **Resume Capability**: Includes logic to handle Google Apps Script 30-minute execution limits by allowing the script to pick up where it left off.

## Prerequisites

Before using this script, ensure you have the following:

1.  **Google Workspace Account**: Access to Google Classroom and Google Drive.
2.  **Google Sheet**: A new Google Sheet to host the script.
3.  **Google Appscript Extension Script Project**: You must enable the **Google Classroom API** and **Google Drive API** in the Apps Script Services menu.

## Setup & Installation

1.  Create a new Google Sheet.
2.  Rename the first tab to `Sheet1`.
3.  Open the Extensions menu and select **Apps Script**.
4.  Copy the contents of `Google_Classroom_Grade_Report_Generator.gs` into the script editor.
5.  **Enable Services**:
    * Click on the `+` icon next to **Services** in the left sidebar.
    * Add **Classroom API**.
    * Add **Drive API**.

### Data Preparation (`Sheet1`)

In the `Sheet1` tab of your spreadsheet, you must provide student and guardian emails in the following columns:

| Column    | Header        | Content                               |

| **G**     | Student Email | The student's school email address.   |
| **H**     | Guardian 1    | Email address of the first guardian.  |
| **I**     | Guardian 2    | Email address of the second guardian. |
| **J**     | Guardian 3    | Email address of the third guardian.  |

**Example Format:**
```text
Column G                    Column H            Column I            Column I
StudentEmail1@School.com    Parent1@Email.com   Parent2@Email.com   Parent3@Email.com
StudentEmail2@School.com    Parent1@Email.com   Parent2@Email.com   Parent3@Email.com
StudentEmail3@School.com    Parent1@Email.com   Parent2@Email.com   Parent3@Email.com
StudentEmail4@School.com    Parent1@Email.com   Parent2@Email.com   Parent3@Email.com

If you have more than 180 students, you'll have to change how many rows it reads in the script.

Configuration
You must modify a few lines in the script to match your specific environment:

IT Email Address:

Locate the sendEmails function.

Replace "yourITemail@email.com" with your actual IT/Admin email address.

School Logo:

Locate the addLogo function.

Replace the URL in the insertImageToCellWithImageBuilder function with the direct download link to your school's logo.

Note: The URL format should look like https://drive.google.com/uc?export=download&id=YOUR_IMAGE_ID.

Usage
To generate the reports, run the functions in the specific order listed below. You can run these from the Apps Script editor toolbar.

ListCoursesAllTabsLocalArray: Fetches course data and builds the initial student sheets.

calculateGradesLocalArray: Calculates the weighted grades for all students.

setColumns: Formats the spreadsheet columns for readability.

addLogo: Inserts the school logo into the report headers.

sendEmails: Generates PDFs and emails them to the addresses listed in Sheet1.

Handling Timeouts
Google Apps Script has a 30-minute execution limit. If you have many courses or students, the script may time out.

If the script stops, check the ListCoursesAllTabsLocalArray settings.

You can adjust variables like startListCoursesFromSpecificCourse and courseNumberToContinueFrom to resume processing from a specific course index.


Disclaimer
This script is provided as-is. Please ensure you test the email functionality with a test address before sending reports to the entire student body.
