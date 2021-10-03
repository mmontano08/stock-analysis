# Stock-Analysis (VBA Challenge)

## Overview of Project:
In this challenge I will edit, or refactor, the module 2 solutions code to loop through all the data. Then I will determine whether refactoring my code successfully made the VBA script run faster. At the end, I will present a written analysis that explains my findings.

### Purpose:
The purpose of this challenge is to use VBA to help us analyze the stocks that we were given. During the beginning of this module we were able to provide Steve's parents with great statistics on 12 different stocks. For this challenge we wanted to see if we changed the arrays, for loops as well as nested loops, it would be able to result in providing a faster output. 

## Results:
After analyzing results, it appears that 2017 had a more positive return and were safe to say that it were safe investments. Particularly, Steve's parents were looking into DQ, which had the greatest return of all. Were as 2018, there was an overall deline on nearly all the stocks. 

![VBAChallenge2017](https://user-images.githubusercontent.com/83566868/135769110-7a614525-332d-4231-b17f-db45eaddd347.png)
![VBAChallenge2018](https://user-images.githubusercontent.com/83566868/135769114-ccf8e124-4001-4be6-8975-8abf6127fe05.png)

The execution times of the original script and the refactored script differences were in fact slower on the second iteration resulting in taking more than 8 seconds (8.98) compared to less than one second in the first iteration.

## Summary:
- The advantages of refactoring code is that the code is fresher, easier to understand or be able to read, it is a lot less complex and easier to maintain. I believe it is an important part of the coding process. 

- Typically there should not be any disadvantages refactoring a code, because the whole purpose of it is to refresh te code by making it simplier to understand.

- When refactoring the original VBA script, having a module to start with, we realized it would become more quickly inefficient. Steve wanted to add additional stocks and that at some point, it will not be possible to input new stock tickers or it would become very time restrictive, adding manually was not going to feasable. Since our results were mixed in that the running of the second or folowing years through the new module are much slower, more coding would need to be executed to address or determine the issue.
