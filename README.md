# Excel/CSV Report Automation

Python automation that cleans Excel/CSV files and generates consolidated business reports.

## Overview

This project demonstrates a practical automation workflow for small businesses that need to process spreadsheet data and generate recurring reports.

The script imports a CSV file, cleans and standardizes the data, calculates business KPIs and exports a consolidated Excel report with multiple sheets.

## Problem

Many businesses still rely on manual spreadsheet work to prepare recurring reports.

Common issues include:

- copying and pasting data manually;
- inconsistent spreadsheet formats;
- repetitive calculations;
- reporting delays;
- human errors;
- lack of standardized outputs.

## Solution

This project automates the reporting workflow using Python.

The automation:

- imports sales data from CSV;
- validates required columns;
- cleans dates and numeric fields;
- calculates revenue and KPIs;
- summarizes performance by category;
- summarizes performance by sales channel;
- creates a monthly report;
- exports everything to a structured Excel workbook.

## Features

- CSV data import
- Data validation
- Data cleaning
- KPI calculation
- Revenue analysis
- Category report
- Sales channel report
- Monthly report
- Excel export with multiple sheets

## Example KPIs

- Total revenue
- Total orders
- Total units sold
- Average ticket
- Revenue by category
- Revenue by sales channel
- Monthly revenue

## Tech Stack

- Python
- Pandas
- OpenPyXL
- Excel/CSV

## Project Structure

```text
excel-report-automation/
├── README.md
├── requirements.txt
├── .gitignore
├── src/
│   └── main.py
├── data_sample/
│   └── sample_sales_data.csv
├── output/
│   └── monthly_sales_report.xlsx
└── screenshots/
