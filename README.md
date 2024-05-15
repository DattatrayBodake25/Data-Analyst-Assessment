# Data Analyst Intern Assessment

## Introduction

This repository contains the code and resources for the Data Analyst Intern Assessment. The assessment task involves completing a function to save extracted tables in Excel or CSV format from provided HTML files.
It's recommended to properly analyze the html file to understand what data needs to be saved. Ensure to sanitize all the inputs and drop any null values to maintain a coherent database structure.

The goal is to extract the data such that it can be used for data visualization purposes including:

- Single entity visualization
- Multiple entity relationship visualization (how two entities are related, like two people working in the same department during the same time)

## Prerequisites

Ensure you have the following dependencies installed:

- Python (>=3.6)
- Selenium
- BeautifulSoup
- Pandas
- pg8000
- psycopg2
- webdriver-manager

You can install these dependencies using pip:

```bash
pip install selenium beautifulsoup4 pandas pg8000 psycopg2 webdriver-manager
```

## Setup

1. Clone this repository to your local machine.

```bash
git clone <repository_url>
```

2. Install the required dependencies as mentioned above.

## Usage

1. Navigate to the project directory.

```bash
cd Data_Analyst_Intern_Assessment
```

2. Open the Python script containing the code.

```python
supremo.py
```

3. Complete the TODO section in the script. You need to create a new function to save all the extracted tables in Excel format.
4. Run the script.

```bash
python supremo.py
```

5. Verify that the tables are saved in Excel format in the `tables` folder.

## Files

- `supremo.py`: Contains the main Python script for the assessment.
- `tables/`: Directory containing sample HTML files of extracted tables.
- `README.md`: This file, providing instructions and information about the assessment.

## Contributing

Feel free to contribute to this repository by forking it and creating pull requests with your changes.

## Submission

Once complete drop a mail to hi@lamarr.tech with the Github link.
