# Traceability Matrix Export

This repository provides a Laravel export class using [Laravel Excel (Maatwebsite)](https://laravel-excel.com/) to generate a traceability matrix in Excel format.

## Features

- Exports traceability data into a structured Excel file.
- Includes headings for each traceability level (User Needs, Design Inputs, Outputs, Verification, Validation).
- Auto-sizes columns for better readability.
- Applies border styling:
  - **Medium borders** for headings.
  - **Thin borders** for data rows.

### Requirements

- Laravel 9+ (or compatible version)
- [`maatwebsite/excel`](https://github.com/Maatwebsite/Laravel-Excel) package
