# Radisson Registration

Radisson Registration is a hotel staff registration dashboard built with React, TypeScript, and Vite. It is designed for HR or people-operations teams to register daily employee arrivals, review department totals, manage the employee directory, and export daily reports.

## Features

- Secure sign-in screen for the registration desk demo flow
- Daily employee registration form with duplicate prevention for the same date
- Employee search and quick fill from the local employee database
- Dashboard with total employee counts, daily registrations, and target progress
- Department breakdown charts for the full directory and the selected day
- Daily report table with the ability to remove registrations
- Export of daily reports to Excel and PDF
- Employee directory import from Excel or CSV files
- Settings for theme mode, officer identity, hotel name, and daily target
- Local persistence with `localStorage` so data stays on the same browser

## Tech Stack

- React 19
- TypeScript
- Vite
- `xlsx` for spreadsheet import and report generation

## Project Structure

- `Reg_Project/src/App.tsx` contains the main application logic and UI
- `Reg_Project/src/App.css` contains the styling for the dashboard and login view
- `Reg_Project/src/main.tsx` bootstraps the React app
- `Reg_Project/index.html` is the Vite entry point

## Getting Started

### Prerequisites

- Node.js 18 or newer
- npm

### Install Dependencies

```bash
cd Reg_Project
npm install
```

### Run the App Locally

```bash
npm run dev
```

Vite will print a local URL in the terminal. Open it in your browser to use the app.

### Build for Production

```bash
npm run build
```

### Preview the Production Build

```bash
npm run preview
```

### Lint the Code

```bash
npm run lint
```

## How It Works

1. Sign in with any access code in demo mode.
2. Register staff members from the Registration page.
3. Search the employee directory or enter details manually.
4. Switch the report date to review past daily entries.
5. Export the selected day as Excel or PDF when needed.
6. Open Settings to update the theme, officer profile, and daily target.

## Data Storage

The app stores registrations, employee data, preferences, and session state in the browser using `localStorage`. Clearing browser storage will reset the app to an empty state.

## Notes

- Employee imports accept `.xlsx`, `.xls`, and `.csv` files.
- The daily report will prevent duplicate registration entries for the same employee on the same date.