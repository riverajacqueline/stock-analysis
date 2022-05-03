# Stock Analysis with VBA

## Overview of Project

In this project, VBA will be used to:
  - Create Variables
  - Write loops and If-Then statements with operators
  - Use an index to access data in an array
  - Reuse, debug, and comment on code
  - Apply visual, numeric, and conditional formatting
  - Measure code performance

### Purpose
The purpose of this project was to create an analysis for various green energy stocks in the years of 2017 and 2018 to determine whether DAQO was worth investing in. The goal for this code was to be refactored in order to perfrom more efficiently.

## Results

### Analysis of Green Engery Stocks from 2017 and 2018

When comparing the Total Daily Volume and Return for Green Energy Stocks in 2017 and 2018, it can been seen that in 2017, all but one stock had a positive return in relation to their total daily volume. When looking at analysis for 2018, all but two stocks had a **decrease** in totaly daily volume as well as a **negative** return. ENPH and RUN were the only two stocks that were able to generate a positive return for 2018. This shows that an investment into DAQO might not be the best choice for. See images below.

#### 2017 Total Daily Volume and Return
![image](https://user-images.githubusercontent.com/103764279/166294212-ddcb37be-94d4-4bca-9d25-ebdd06818609.png)

#### 2018 Total Daily Volume and Return
![image](https://user-images.githubusercontent.com/103764279/166294660-3d1d5ae6-09a8-4cc4-8c7a-4674c2182300.png)

### Deconstructing the Code
#### The Original Code
The original code that was first used was created with a nested loop. This loop switches back and forth between the "All Stocks Analysis" worksheet and the "2018" stocks worksheet.

![image](https://user-images.githubusercontent.com/103764279/166324954-d170d282-b30a-461b-8787-dbdf0eb107d2.png)
![image](https://user-images.githubusercontent.com/103764279/166324987-0094600f-5372-49f7-afb6-b7fb1e4523a0.png)
![image](https://user-images.githubusercontent.com/103764279/166325052-9e09ba3f-535a-45e1-be68-9b00197f80ab.png)

#### The Refractured Code
The refractored code was created with multiple loops instead of a nested loop. In the first loop, three arrarys were created to store the data while the the second loop output the data into a separate worksheet.

<img width="494" alt="Refractored_Code1" src="https://user-images.githubusercontent.com/103764279/166328134-a0cccc94-a20c-4d0d-8e93-b9c867e72bb3.png">
<img width="662" alt="Refractored_Code2" src="https://user-images.githubusercontent.com/103764279/166328178-23d04c3f-fa92-429c-bc91-ce24cc274a24.png">
<img width="679" alt="Refactored_Code3" src="https://user-images.githubusercontent.com/103764279/166328231-e7824aa6-d13a-4fff-a4a3-83a109f1dd03.png">
<img width="620" alt="Refactored_Code4" src="https://user-images.githubusercontent.com/103764279/166328255-95ef3efb-a76e-4a52-82e9-01c43ea77cfe.png">

#### Run times of the Original Code and the Refactored Code
When looking at the run times of both the original code and the refactored code, the refactored code runs way faster than the original code. By creating arrays in the refactored code, it is able to run faster than the original code because it not having to switch back and forth between worksheets.

##### ***The run times of the Original code are as follows:***

<img width="257" alt="Origninal_Code_RunTime_2017" src="https://user-images.githubusercontent.com/103764279/166328940-529268fb-ccdf-40cb-a691-a232eb2a1db6.png">
<img width="250" alt="Original_Code_RunTime_2018" src="https://user-images.githubusercontent.com/103764279/166328967-86d3b3fe-3bf8-4044-b318-77811a5a6988.png">

##### ***The run times of the Refactored code are as follows:***

<img width="183" alt="Refactored_Code_RunTime_2017" src="https://user-images.githubusercontent.com/103764279/166329467-355d6c35-3824-4080-83e0-7c97539960f0.png">
<img width="182" alt="Refactored_Code_RunTime_2018" src="https://user-images.githubusercontent.com/103764279/166329505-809be0d7-d377-4d08-8a3f-334175bd60e1.png">

## Summary

#### What are the advantage or disadvantages of refactoring code?
##### Advantages
- Refactoring the code helps the program run more efficiently by increasing the speed of the run time.
- Refactoring the code helps made the code easier to read and understand.
- It can help find bugs within the program.
##### Disadvantages
- Refactoring code is time consuming.
- It might be hard to do with a large codes.
- The person who is refactoring the code has to be able to know what the code is asking for so that they can be able to improve it.
- Refactoring can actually create bugs also, which can then lead to altered outcomes.
#### How do these pros and cons apply to refactoring the original VBA script?
- A pro for refactoring the original VBA script is that it led to more efficient code and decreased run time.
- A con is that the code is a little confusing to read if you do not know what is going on in the code.
