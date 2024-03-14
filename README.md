<h1 align="center"> VBA-Challenge </h1> <br>
Utilize VBA scripting to analyze generated stock market data.

## Table of Contents

- [Introduction](#introduction)
- [Getting Started](#getting-started)
- [Build Process](#build-process)
- [Considerations](#considerations)
- [Author](#author)

## Introduction

The purpose of this challenge is to be able to create a script that loops through the provided stocks and output specific information. 
Columns were newly created for the following findings: 
- The ticker symbol.
- The calculated yearly change, based on the sum of opening and closing values.
- The percent in change.
- The total stock volume.


## Getting Started

Prior to execution, you must have Microsoft Excel installed. You must also have all your Macros enabled; this can be done within the Trust Center Settings. 


## Build Process

Create a script that loops through all the stocks for one year and outputs the following information:

- The ticker symbol
- Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
- The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
- The total stock volume of the stock. The result should match the following image:

    ![image](https://github.com/myoingco/VBA-challenge/assets/160566342/73e91c19-09f4-4549-b504-7e926297726e)

Create a second loop to establish the following:

- Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:

    ![image](https://github.com/myoingco/VBA-challenge/assets/160566342/0024593f-f1a4-4a7b-809a-4fa3b8320262)


## Considerations

- Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.
- Make sure to use conditional formatting that will highlight positive change in green and negative change in red.
- Make sure that the script acts the same on every sheet. The joy of VBA is that it takes the tediousness out of repetitive tasks with the click of a button.


## Author

[ Meichel Yoingco](https://github.com/myoingco)


