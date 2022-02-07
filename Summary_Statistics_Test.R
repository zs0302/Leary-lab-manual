### test

### simulation
### Generate the simulated data with 4 independent variables (2 continuous and 2 categorical) and 2 dependent variables (1 continuous and 1 categorical)
Input_Data <- data.frame(x1 = rnorm(100),
                         x2 = as.factor(paste0("Text", rbinom(100, 6, 0.5))),
                         x3 = rgamma(100, 3),
                         x4 = as.factor(rbinom(100, 1, 0.5)),
                         y1 = rnorm(100),
                         y2 = as.factor(c("Yes", "No")[rbinom(100, 1, 0.5)+1]))

### Set Output Path
Output_Path <- choose.dir()

### Run the main function
### Note: the data should be pre-processed
### 1. All categorical variables should be saved as factors; the function, as.factor(), can be used in the pre-processing.
### 2. The reference level should be the first level; the function, relevel(), can be used for this purpose.
### 3. there should be no missing data
Summary_statistics(Input_Data = Input_Data,
                   Independent_Vars = c("x1", "x2", "x3", "x4"),
                   Dependent_Vars = c("y1", "y2"),
                   Output_Path = Output_Path)

