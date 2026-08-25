#  FinAccrual Excel Add-in (Frontend)

This directory contains the Microsoft Excel Taskpane Add-in built using Office JS. It allows users to authenticate with QuickBooks or Xero, create worksheets for accrual tracking inside Excel, and pull metadata directly into spreadsheet cells.

##  Key Features

- **Multi-provider Support**: Dropdown configuration selectors for QuickBooks and Xero.
- **Accrual Sheet Setup**: Automation script that sets up standard columns across 5 section blocks (Company, Accounts, Classes, Locations, Entities) with styled headers and freeze panes.
- **Data Synchronization & Streamed Pull**: Pull Master Data stream integration that populates master data in real-time.
- **Incremental Refresh Schedule**:
  - Requires **Setup Master & Input Sheets** AND **Pull Master Data** steps to be completed before allowing execution.
  - Compares fetched live records against existing Excel IDs across all 5 table blocks to prevent duplicate rows.
  - Appends only new or updated records below existing sheet data.
  - Displays accurate notifications reporting exact updated record counts (e.g. `1 updated record added.` or `0 updated records added.`).
- **Free Trial Lifecycle & Expiration Protection**:
  - 2-minute free trial lifecycle with real-time expiration monitoring.
  - Clears sheet data upon trial expiration and presents the Upgrade modal.
  - Back-arrow navigation guard on the Plans View (`#btnPlansBack`) that prevents returning to the Dashboard when trial is expired.
  - Strict data action guards blocking Setup, Pull, and Refresh when trial is expired.
- **Company Switching**:
  - Smooth active company selection with instant state cleanup.
  - Automatically resets step completion flags and clears pagination cursors upon company switch.

##  Folder Structure

```
Frontend/
├── assets/             # Icons and images for manifest
├── src/
│   ├── shared/         # Shared API error handling and notification utilities
│   └── taskpane/
│       ├── controllers/# App, view, data action, and trial controllers
│       ├── router/     # ViewRouter navigation engine
│       ├── services/   # ApiService, ExcelService, DashboardService, AuthService
│       ├── state/      # AppState store
│       ├── styles/     # Taskpane CSS styles
│       ├── taskpane.css
│       ├── taskpane.html  # Main Taskpane UI layout
│       └── taskpane.js    # Entry point & module initializer
├── manifest.xml        # Excel Office Add-in Manifest
└── webpack.config.js   # Webpack bundler configuration
```

##  Getting Started & Sideloading

### Prerequisites
- Node.js (v18.x+)
- Desktop Excel (Windows/macOS) or Excel Online subscription

### Installation

1. Navigate to this directory:
   ```bash
   cd Frontend
   ```
2. Install npm packages:
   ```bash
   npm install
   ```

### Running & Building

1. Start the dev server and sideload the add-in inside desktop Excel:
   ```bash
   npm start
   ```
   *This starts the server on `https://localhost:3000` and automatically opens a new Excel workbook containing the FinAccrual add-in.*

2. Build for production:
   ```bash
   npm run build
   ```

3. To stop the add-in and clear sideloading state:
   ``bash
   npm run stop
   ```

4. To validate the manifest file:
   ``bash
   npm run validate
   ```
