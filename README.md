# Projecting-for-Algriculture
A programm prepared for the Mathematical Contest in Modeling
## Brief Introdution:
this programm only provide the code for the problem one
## Introduction for the functions
### 1) As for taking in the formats in .xslx file to put in data
You should define the path accroding to your actual cases
The `pandas` library would be needed to read the files
### 2) As for putting the data out to a new created file
#### 1） Information needed
You should install the `openpyxl` library to use the aiming function
And you should also choose the location where you want to save your created file by your own
The amount of years needing calculation should also be defined by users
#### 2) Mechanism for outputting
The formats would be filling along with the loops, each loop would fill one format
During each case, an object for calculating the relative data would be created, and the formats would be filled accroding
to the components after the working process of the objects
### 3)For the related materials
The introduction about the problem and the format files which include the original information are all provided in the file blank
### 4)Brief telling about the mechanism for the provided class
1)Land blocks and the types of crops would be take consideration as different types of objects respectively
2)For each type of crops class, its objects would record the most benificial method to till the land
3)A priority for the crop types is constructed, prove that the most benificial way would be taken in consideration first
4)Each time a `crop` object is chosen, a function for operating the project would be activated for planting utill finished
5) During the process 4 there would be a telling function to prove the crops are planted accroding to the rules
6) After allocating land for the unqueued objects, the priority queue would be updated in order for the next stage
`The changes are made accroding to the changes of most benificial planting method`
7)After the first season, the next loop for planting would be activated for the season 2, and the priority queue would be modulated for this period
8)The results would be stored in the related attributes for updating the original data and output the results


