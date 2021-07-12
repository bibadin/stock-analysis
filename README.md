# stock-analysis
## Overview of Project
The purpose of this project was to analyze thousands of stocks to determine which year produced a better yearly return. This project was also completed with refactored code to be more concise and efficient. 

## Results
-11 stocks in 2017 grew between 5.5-199.4% while majority of stocks in 2018 showed decreased return. 

  -Original script ran 1.8125 seconds for 2017 and 1.8515 seconds for 2018. 

Refactored script 
![VBA_Challenge_2017.png](path/to/VBA_Challenge_2017.png)
![VBA_Challenge_2018.png](path/to/VBA_Challenge_2018.png)

  -The refactored code ran much faster compared to original script

    -Both the original and refactored script for 2018 ran more slowly compared to 2017.

Example of code: If Cells(I, 3) > 0 Then
       
        Cells(I, 3).Interior.Color = vbGreen
        
    ElseIf Cells(I, 3) < 0 Then
        
        Cells(I, 3).Interior.Color = vbRed
    Else
        
        Cells(I, 3).Interior.Color = xlNone
        
        End If

Example of refactored code: If Cells(I, 3) > 0 Then
            
            Cells(I, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(I, 3).Interior.Color = vbRed
            
        End If
        
  -The refactored script is obviously easier to read which is helpful when multiple people are working with the same script.

## Summary
### Advantages

-Advantages to refactoring code in general is that it facilitates a more efficient and quicker process. The code overall is clearer and easier to read.

### Disadvantages

-Disadvantage of refactoring code is that increase in errors could occur especially if different people are apart of the refactoring. Also it can take a longer period of time to condense the code and run properly.

### Pros
-Improved the code's readability
- Decreased complexity
- Faster running code that leads to improved performance

### Cons
-New problems could be introduced when trying to condense the code
