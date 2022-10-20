# VBA of Wall Street

## Overview of Project

### Purpose

The purpose of this project was to refactor the VBA code written during this module to make it run faster. While the original code for the green_stocks Excel spreadsheet, renamed [VBA_Challenge](https://github.com/stwpf01/Stock_Analysis/blob/main/VBA_Challenge.xlsm), ran quickly enough, it would take considerably longer to execute if it where to analyze more data. Thus, refactoring the code to be more efficient was necessary. The results of the refactored code will be detailed below.

## Results

The refactored code ran substantially faster than the original code. To show the difference, here is an image of a Message Box with the time it took to execute the code to analyze the 2018 stocks originally: 

![Original Code 2018](https://github.com/stwpf01/Stock_Analysis/blob/main/Resources/VBA_Challenge_2018_NonRefactored.png) 

And here is another Message Box showing how long it took to execute the refactored code for the same year:

![Refactored Code 2018](https://github.com/stwpf01/Stock_Analysis/blob/main/Resources/VBA_Challenge_2018.png)]

Similar results happened when executing the original and refactored code for the 2017 stocks. Example images can be found in the [Resources](https://github.com/stwpf01/Stock_Analysis/tree/main/Resources) folder.



## Summary

If it wasn't made clear with the images showing the difference in time execution for the codes, refactoring code greatly helps with making the code run faster by cleaning it up and combining multiples lines of code into single lines. The disadvantage of refactoring code is the possibility that the attempt of making the code run better makes it not run at all. This does, however, reinforce the notion of making copies of ones code to make sure there is always a working copy to fall back on when experimentation does not yield desired results.

For this VBA code specifically, the shortening of code and making it run faster is a clear advantage of refactoring. Disadvantages, beyond not saving original code that is, really seems to stem from the coders limitations on their knowledge of the coding language they are refactoring. While there are aspects of the refactored code I understand, parts of it still seem above me. This isn't necessarily a bad thing as everything has a learning curve to it. If anything it means one needs to keep at the material so as time goes on a deeper understanding is obtained. 