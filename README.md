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

Refactoring code has several advantages. It can help to reduce complexity and duplication in the code, especially for legacy code. It also improves the organization and efficiency of the code to achieve better performance and increase its overall readability and maintability over time. Another key benefit is that it makes the code more extensible and flexible. Moreover, it can help other developers understand the code better. It can be beneficial in situations where there are multiple developers collaborating on a project but on a different aspects of the code; refactoring helps developers to understand the end to end logic and flow of the code if the code is well organized and clean.

There are some disadvantages with refactoring code. Refactoring can introduce defects but this can be mitigated with proper peer review and testing. It can be time consuming, especially for a large or legacy code base. Another disadvantage is that refactoring does not address any underlying architectural issues. Lack of or insufficient test cases can also add risks to refactoring.

## Advantages and Disadvantages of the original and refactored VBA script

As a result of refactoring the original VBA script, the run time performance improved from over 1 second to less than almost one fourth the time.

Attached below are the screenshots that indicate the run time for our new analysis.
<img width="1440" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/80140082/112773639-8d8f5700-8feb-11eb-92c4-c565e7ababf8.png">

<img width="1440" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/80140082/112773592-6cc70180-8feb-11eb-98ec-305558a02c67.png">

![image](https://user-images.githubusercontent.com/80140082/112773684-ba436e80-8feb-11eb-96e6-743bcfb253d8.png)

![image](https://user-images.githubusercontent.com/80140082/112773715-e5c65900-8feb-11eb-870c-5e26c3116a37.png)


