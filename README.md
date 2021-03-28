# Stock-Analysis

# Overview of Project

## Purpose

The purpose of this project was to refactor Microsoft Excel VBA code to collect total daily volume and return for 12 green energy stock prices for the years 2017 and 2018 and determine if these stocks are worthwhile investing in. The starting point was to reuse code to set up a similar structure. The goal of the challenge was to reuse and refactor the original code to perform the analysis for 2017 and 2018 stocks and improve the overall efficiency of the code.

## Data

The data used in this analysis were two separate worksheets for 2017 and 2018 stock information. The data contained therein included the ticker index, the stock issue date, the opening price, the high and low price, the adjusted closing price, and the volume for each stock. The objective of the analysis was to retrieve the ticker, the total daily volume, and the return on each stock.

The challenge instructions were to add code to determine the total daily volume and return for each stock for each year and refactor the code structure.

## Results

Before refactoring the code, I reviewed the code that I could reuse and copied the code to create the table headers,the input box for the year that I want to run the analysis, the ticker array, and to format the header rows and conditional color formatting. I also made sure that I activated the correct worksheet to output the data.

The instruction and code as written in the file are below.


# Summary

## Advantages and Disadvantages of Refactoring Code

Refactoring code has several advantages:
    ### 1) It can help to reduce complexity and duplication in the code, especially legacy code
    ### 2) It improves the organization and efficiency of the code to achieve better performance and increase its overall readability and maintability over time.
    ### 3) It makes the code more extensible
    ### 4) It can help other developers understand the code better. It could be beneficial in situations where there are multiple developers collaborating on a      project but on a different aspects of the code; refactoring helps developers to understand the end to end logic and flow of the code if the code is well organized and clean.

There are some disadvantages with refactoring code. They can be:
    1) It may introduce defects but this can be mitigated with proper peer review and testing
    2) It can be time consuming, especially for a large or legacy code base

## Advantages and Disadvantages of the original and refactored VBA script

Overview of Project: Explain the purpose of this analysis.
Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?
