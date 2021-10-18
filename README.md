# wk2_VBA_challenge

## Overview of Project
### Background
    We have created an analysis during the module for a particular stock "DQ".

### Purpose
    The purpose of this analysis is to expand the dataset to include the entire stock market over the last few years. 

    We learned to _refactor_the code to loop through all the data one time to collect information. We will determine if refactoring the code successfully has made the VBA script run faster.

## Results 
### Stock performance
    The overall performance of stocks in 2017 is much better than in 2018. 2 stocks "RUN" and "ENPH" are exceptions where they performed better in 2018 than the year before.

### Execution time
    The run time after using refactored code. 
    Instead of looping thru the entire dataset times the number of tickers as we did originally, we now loop thru the dataset one time, inside the the loop, we look for the starting and ending prices to complete the calculation. So we have dramatically reduced the number of steps, using less memory to make the code more efficient.

## Summary
### About refactoring in general
#### Pros
    - High chance of code enhancement
    - Bug fixing
    - Peer review
#### Cons
    - Refactoring takes time
    - Cost of refactoring is higher than rewriting from scratch

### about refactored VBA script
#### Pros
    - run time is much faster
    - better understanding of VBA by more practising
    - 
#### Cons
    - time consuming
    - error prone 
