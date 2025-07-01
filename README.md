# Weather Analysis Project

This project analyzes a large weather dataset containing 1,000,000 rows of data, focusing on temperature, humidity, precipitation, and wind speed across multiple cities. The analysis is performed using Python with the Pandas, NumPy, Matplotlib, OpenPyXl libraries, with results visualized in graphs and saved for further review.

## Features
- Loads and processes a 1,000,000-row weather dataset.
- Filters data by city and date, including specific months.
- Calculates statistical measures (e.g., mean temperature).
- Generates plots for temperature, humidity, wind speed, and precipitation over time.
- Organizes output files in a dedicated `Analyzed` folder.

## Skills Demonstrated
- Python
- Pandas (data manipulation and analysis)
- NumPy (numerical computations)
- Matplotlib (data visualization)
- OpenPyXl
- File handling and GitHub usage

## Project Structure
- `main.py`: Main script for data analysis and plotting.
- `weather_data.csv`: Input dataset with 1,000,000 rows.
- `Analyzed/`: Folder containing generated plots and Excel tables.

## How to Run
1. Ensure Python is installed with required libraries: `pip install pandas numpy matplotlib`.
2. Place the `weather_data.csv` file in the project directory.
3. Run the script: `python main.py`.
4. Check the `weather_analysis` folder for generated plots.

## Notes
- The script filters data for a specific city (e.g., San Diego) and month (e.g., July 2024).
- Plots are saved individually for each date, showing multiple weather parameters.
- Empty folders are handled, and the project is optimized for large datasets.