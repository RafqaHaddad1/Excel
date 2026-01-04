# Excel Import & Viewer

A simple ASP.NET Core MVC project that lets you **upload an Excel file** (`.xls` / `.xlsx`), **read its sheets and rows**, and **display the extracted data** in the UI using **ExcelDataReader**.

## Features

- Upload Excel files from the browser
- Read Excel content using **ExcelDataReader**
- Support for multiple sheets (data grouped by sheet)
- Basic validation (file selected, allowed extensions)
- Displays parsed output in a “Success” view (or similar)


## Tech Stack

- **ASP.NET Core MVC**
- **C#**
- **ExcelDataReader**
- Razor Views


## Getting Started

### Prerequisites

- **.NET SDK** (matching the project target framework)
- Visual Studio / VS Code

### Install & Run

1. Clone the repository:
   ```bash
   git clone https://github.com/RafqaHaddad1/Excel.git
   cd Excel
