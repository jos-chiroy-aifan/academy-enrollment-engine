# Academy Enrollment Engine

Event-driven enrollment and scheduling engine built on Google Apps Script. This project automates administrative tasks for an academy, managing student enrollments, professor schedules, financial reporting, and form synchronization.

## Features

- **Dynamic Form Synchronization:** Automatically updates Google Forms based on programs, schedules, and availability.
- **Automated ID Generation:** Generates unique identifiers for students, professors, programs, and enrollments.
- **Scheduling & Availability Engine:** Enforces physical room capacity, manages professor availability, and builds schedule dashboards.
- **Automated Notifications & Emails:** Sends customized emails to students and professors for registration confirmation, new student assignments, schedules, and financial reports.
- **Financial & Payment Tracking:** Generates monthly payment records and weekly financial reports.
- **Migration Utilities:** Helps migrate old form responses and handle retries for student/schedule migrations.

## Tech Stack

- Google Apps Script
- JavaScript / Node.js (for local typing and dependency management)

## Setup

1. Install dependencies:
   ```bash
   npm install
   ```
2. Deploy the scripts using `clasp` (Command Line Apps Script Projects) to your Google Apps Script environment.

## Project Structure

- `src/` - Contains all Google Apps Script source files.
  - Utilities and core components
  - ID Generators
  - Form Handlers and Sync scripts
  - Dashboards and Schedule builders
  - Automated Email services
- `package.json` - Node project configuration and dependencies.

## License

ISC
