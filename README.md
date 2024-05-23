# Financial Statements Analysis Tool

![Financial Statements Analysis Tool](assets/financial-statements-analyser.png)

Visit the deployed site: [Financial Statements Analysis Tool](https://fin-statements-analysis-tool-13d33c22c8d1.herokuapp.com/)

This tool updates and analyses the financial performance of a small company. It uses Google Sheets and Python to manage data extracted from financial statements, calculates financial ratios, and carries out benchmark and trend analyses.
## Table of Contents

- [Project Description](#project-description)
  - [User goals](#user-goals)
  - [Site owner goals](#site-owner-goals)
- [Pre-development](#pre-development)
- [Development](#development)
- [Features](#features)
  - [Update Profit and Loss Account](#update-profit-and-loss-account)
  - [Update Balance Sheet](#update-balance-sheet)
  - [Generate Financial Statements](#generate-financial-statements)
  - [Calculate Financial Ratios](#calculate-financial-ratios)
  - [Update Ratios in Google Sheets](#update-ratios-in-google-sheets)
  - [Options for users to select next step](#options-for-users-to-select-next-step)
  - [Analyse and Compare Ratios](#analyse-and-compare-ratios)
  - [Benchmark Analysis](#benchmark-analysis)
  - [Trend Analysis](#trend-analysis)
  - [Summary](#summary)
  - [Future Implementations](#future-implementations)
- [Accessibility](#accessibility)
- [Technologies Used](#technologies-used)
- [Resources](#resources)
  - [Libraries](#libraries)
- [Testing](#testing)
- [Future Updates](#future-updates)
- [Validation](#validation)
- [Deployment](#deployment)
- [Heroku](#heroku)
- [Bugs](#bugs)
- [Credits](#credits)

## Project Description

This tool updates and analyses the financial performance of a small company. It manages financial data, calculates various financial ratios, and carries out benchmark and trend analyses. The data is stored in Google Sheets and Python is used to update and analyse it.

The main goals of this project are to:

- Reduce the time spent updating financial statements manually.
- Automate the calculation of financial ratios.
- Provide a comprehensive financial report.
- Compare financial performance against industry benchmarks and historical data.

### User goals:

- Update financial statements quickly and accurately.
- Automatically calculate and analyse financial ratios.
- Compare current financial performance with historical data and industry benchmarks.

### Site owner goals

- Provide an easy-to-use tool for financial analysis.
- Ensure data accuracy and reliability.

## Features

### Update Profit and Loss Account

The user is prompted to update the following accounts: sales revenue, inventory, rent, and interest expenses. A specific number range is provided for users. 

### Update Balance Sheet

A user is prompted to update the following accounts: cash and cash equivalents and short-term loans. The specific number range is provided for users. 

Once the data is entered, the profit and loss and balance sheet tabs are updated in Google Sheets.

### Generate Financial Statements

A user is given a choice whether they want to generate updated financial statements or not. 

### Calculate Financial Ratios


The following ratios are calculated automatically:

 - Liquidity ratios: current and quick ratios
 - Profitability ratios: net profit margin and return on assets
 - Solvency ratios: debt-to-equity ratio and interest cover

### Update Ratios in Google Sheets

Once all ratios are calculated, the results are being updated in Google Sheets. Namely, the results are added to the fourth quarter the Ratios Historical Data tab. These Quarter 4 data is after used to carry out various types of analysis as part of the financial reporting.


### Options for users to select next step

After updating the financial data, users can choose what they want to see next. The tool offers the following options:
- financial ratios analysis
- benchmark comparison -
- trend analysis
- run a complete financial report that combines all options above. 

This flexibility allows users to tailor the analysis to their specific needs and interests.


### Analyse and Compare Ratios


The tool provides a detailed analysis of financial ratios, including liquidity, profitability, and solvency ratios. This analysis helps users understand the financial health of the company, highlighting areas of strength and potential concern


### Benchmark Analysis

The tool compares the calculated financial ratios with industry benchmarks. This comparison helps users understand how their company is performing relative to industry standards, identifying areas where the company excels or may need improvement.

### Trend Analysis

The trend analysis feature evaluates financial ratios over four quarters to identify patterns and changes over time. This analysis helps users understand the direction of the company's financial performance, spotting trends that may indicate future challenges or opportunities.


### Summary


The summary provides a short explanation of what is included in the financial report and why these features are important.

#### Future Implementations

1. Add a more detailed ratio analysis.
2. Implement visualisation options for financial data.
3. Enable exporting of the report to PDF and Excel formats.