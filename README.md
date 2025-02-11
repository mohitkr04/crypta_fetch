# Cryptocurrency Data Fetcher and Analyzer

This project fetches live cryptocurrency data for the top 50 cryptocurrencies from the CoinGecko API, performs basic analysis, and presents the data in a live-updating Excel sheet.

## Table of Contents

*   [Project Overview](#project-overview)
*   [Features](#features)
*   [Requirements](#requirements)
*   [Installation and Setup](#installation-and-setup)
*   [Running the Script](#running-the-script)
*   [Output](#output)
*   [Data Analysis](#data-analysis)
*   [Excel Sheet](#excel-sheet)
*   [Contributing](#contributing)

## Project Overview

This Python script automates the process of gathering and analyzing cryptocurrency market data. It leverages the CoinGecko API to obtain real-time information, providing a snapshot of the top 50 cryptocurrencies ranked by market capitalization. The script is designed to be continuously run, updating an Excel spreadsheet with the latest data every 5 minutes.  This provides a live view of market trends and key metrics.

## Features

*   **Live Data Fetching:** Fetches real-time cryptocurrency data from the CoinGecko API.
*   **Top 50 Cryptocurrencies:** Focuses on the top 50 cryptocurrencies by market capitalization.
*   **Data Analysis:** Performs basic analysis, including:
    *   Identifying the top 5 cryptocurrencies by market cap.
    *   Calculating the average price of the top 50.
    *   Determining the highest and lowest 24-hour price changes.
*   **Live-Updating Excel Sheet:**  Automatically updates an Excel spreadsheet with the fetched data and analysis results every 5 minutes.
*   **Formatted Excel Output:** Presents the data in Excel with clear formatting:
    *   Bold headers with a blue background.
    *   Number formatting (commas, decimal places).
    *   Percentage formatting.
    *   Auto-fitting column widths.
    *   Center alignment.
    *   Cell borders.
*   **Error Handling:** Includes error handling for API request failures.
*   **Configuration via .env:** Uses a `.env` file for easy configuration of file paths and update interval.

## Requirements

*   Python 3.7+
*   `requests` library
*   `pandas` library
*   `openpyxl` library
*   `python-dotenv` library

## Installation and Setup

1.  **Clone the Repository (Optional):** If the code is hosted on a platform like GitHub, clone the repository to your local machine:

    ```bash
    git clone https://github.com/mohitkr04/crypta_fetch
    cd crypta_fetch
    ```
    If you have the code directly, simply create a new directory for the project.

2.  **Install Python:** If you don't have Python 3.7 or later installed, download and install it from [https://www.python.org/downloads/](https://www.python.org/downloads/).  Ensure you select the option to "Add Python to PATH" during installation.

3.  **Install Dependencies:** Open a terminal or command prompt *in the project directory* and install the required libraries using `pip`:

    ```bash
    pip install requests pandas openpyxl python-dotenv
    ```

4.  **Create a `.env` File:** Create a new file named `.env` (no file extension) in the *same directory* as your Python script. Add the following lines to the `.env` file, replacing the placeholders with your actual values:

    ```
    EXCEL_FILE_PATH=C:\Users\mohit\crypto_fetch\crypto_data.xlsx
    EXCEL_FILE_NAME=crypto_data
    SLEEP_TIME=300
    ```

    *   `EXCEL_FILE_PATH`:  The *full path* to the directory where you want the Excel file to be created (e.g., `C:\\Users\\YourName\\Documents\\CryptoProject` on Windows, `/Users/YourName/Documents/CryptoProject` on macOS/Linux).  **Use double backslashes (`\\`) on Windows.**
    *   `EXCEL_FILE_NAME`: The desired name for your Excel file *without* the `.xlsx` extension (e.g., `crypto_data`).
    *   `SLEEP_TIME`:  The update interval in seconds (300 seconds = 5 minutes).  Keep this at `300`.

    **Important:** Ensure there are *no spaces* around the `=` signs in the `.env` file.

5.  **Create an Empty Excel File:** In the folder you specified for `EXCEL_FILE_PATH`, create a new, *empty* Excel file with the name you provided for `EXCEL_FILE_NAME` and the `.xlsx` extension (e.g., `crypto_data.xlsx`).  The Python script will populate this file.

## Running the Script

1.  **Open a Terminal:** Open a terminal or command prompt *in the project directory* (where your Python script and `.env` file are located).

2.  **Run the Script:** Execute the Python script using:

    ```bash
    python crypto_fetcher.py
    ```

The script will:

*   Fetch data from the CoinGecko API.
*   Create a Pandas DataFrame.
*   Create (if it doesn't exist) or update the Excel sheet with the data and formatting.
*   Print the data analysis results to the console.
*   Wait for 5 minutes (300 seconds).
*   Repeat the process indefinitely.

To stop the script, press `Ctrl+C` in the terminal.

## Output

*   **Console Output:** Each time the script fetches data, it will print the following to the console:
    *   Confirmation that the Excel file has been updated.
    *   The top 5 cryptocurrencies by market capitalization.
    *   The average price of the top 50 cryptocurrencies.
    *   The cryptocurrencies with the highest and lowest 24-hour price changes.
    *   Any error messages (if there are problems fetching data).

*   **Excel File:** The Excel file (e.g., `crypto_data.xlsx`) will contain a sheet named "Crypto Data" with the following:
    *   **Headers:** Cryptocurrency Name, Symbol, Current Price (USD), Market Capitalization, 24-hour Trading Volume, Price Change (24h, %).
    *   **Data:**  The fetched data for the top 50 cryptocurrencies, updated every 5 minutes.
    *   **Formatting:**  The data will be formatted for readability, as described in the "Features" section.

## Data Analysis

The script performs the following data analysis:

*   **Top 5 Cryptocurrencies by Market Cap:** Identifies and displays the top 5 cryptocurrencies with the highest market capitalization.
*   **Average Price:** Calculates and displays the average current price of the top 50 cryptocurrencies.
*   **Highest/Lowest Price Change:**  Identifies and displays the cryptocurrencies with the highest and lowest 24-hour percentage price changes.

## Excel Sheet

The generated Excel sheet provides a live-updating view of the cryptocurrency market data.  The formatting enhances readability, making it easy to track key metrics.  The sheet is designed to be viewed in a spreadsheet program like Microsoft Excel, Google Sheets, or LibreOffice Calc.  To share a live view, upload the Excel file to a cloud service like Google Sheets or OneDrive and share a view-only link.  Remember to periodically re-upload the local Excel file to the cloud service to keep the online version updated.

## Contributing

Contributions are welcome! If you find any bugs or have suggestions for improvements, please feel free to open an issue or submit a pull request.
