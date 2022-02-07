#===========================================================================================================#

#===================================================== Main Function =======================================#

#===========================================================================================================#

Summary_statistics <- function(Input_Data, Independent_Vars, Dependent_Vars, Output_Path, Description_Only = FALSE){
  library("openxlsx")
  # create work book
  wb <- createWorkbook()

  # data description
  table1_description(Input_Data, Independent_Vars, Dependent_Vars, wb)
  
  if(Description_Only == FALSE){
    # summary statistics
    for(i in Dependent_Vars){
      if(is.factor(Input_Data[,i])){
        table2_categorical(Input_Data, Independent_Vars, i, wb)
      }
      else{
        table3_continuous(Input_Data, Independent_Vars, i, wb)
      }
    }
  }

  
  # Generate Outputs
  saveWorkbook(wb, 
               file = paste0(Output_Path,"\\Summary_Statistics_",
                             gsub("-","",Sys.Date()),
                             ".xlsx"), 
               overwrite = TRUE)
  
  # clear cache
  file.remove(list.files()[which(grepl('.png$',list.files()))])
  
}

#===================================================== Sub Functions =======================================#

# data description
table1_description <- function(Input_Data, ind_var_list, dep_var_list, wb){
  Cate_output <- NULL
  Cont_output <- NULL
  for(i in c(ind_var_list,dep_var_list)){
    # categorical
    if(is.factor(Input_Data[,i])){
      Cate_out_unit <- t(rbind(c(i, rep(NA, length(table(Input_Data[,i]))-1)), 
                               names(table(Input_Data[,i])),
                               table(Input_Data[,i]),
                               round(prop.table(table(Input_Data[,i]))*100,2)))
      Cate_output <- rbind(Cate_output, Cate_out_unit)
    }
    # continuous
    else{
      Cont_out_unit <- c(i,
                         sum(!is.na(Input_Data[,i])),
                         round(mean(Input_Data[,i], na.rm = TRUE),2),
                         round(sd(Input_Data[,i], na.rm = TRUE),2),
                         round(median(Input_Data[,i], na.rm = TRUE),2),
                         paste0("[",round(quantile(Input_Data[,i], na.rm = TRUE)[2],2),", ",
                                round(quantile(Input_Data[,i], na.rm = TRUE)[4],2), "]")
      )
      Cont_output <- rbind(Cont_output,Cont_out_unit)
    }
  }


  # generate output
  addWorksheet(wb,"Data_Description")
  Cont_start_row <- 1
  # adjust column names
  if(!is.null(Cate_output)){
    colnames(Cate_output) <- c("Variables", "Levels", "Frequency", "Percentage")
    writeData(wb,"Data_Description", 
              startCol = 1,
              startRow = 1, 
              x = "Categorical Variables")
    writeData(wb,"Data_Description", 
              startCol = 1, 
              startRow = 2, 
              x = Cate_output)
    Cont_start_row <- nrow(Cate_output)+5
  }

  if(!is.null(Cont_output)){
    colnames(Cont_output) <- c("Variables", "Count", "Mean", "Std", "Median", "IQR")
    writeData(wb,"Data_Description", 
              startCol = 1,
              startRow = Cont_start_row, 
              x = "Continuous Variables")
    writeData(wb,"Data_Description", 
              startCol = 1, 
              startRow = Cont_start_row + 1, 
              x = Cont_output)
  }
}

# odds ratio for 2x2 matrix
unit_odds_ratio <- function(input_matrix){
  out_odds_ratio <- c(1,0,0)
  input_matrix[which(input_matrix==0)] <- 0.5
  out_indicator <- (input_matrix[1,1]==input_matrix[1,2])*(input_matrix[2,1]==input_matrix[2,2])
  out_indicator2 <- (input_matrix[1,1]==input_matrix[2,1])*(input_matrix[1,2]==input_matrix[2,2])
  if(max(out_indicator,out_indicator2)==1){
    return(out_odds_ratio) 
  }
  else{
    output_ratio <- input_matrix[,1]/input_matrix[,2]
    output_odds1 <- output_ratio[1]/output_ratio[2]
    output_std <- sqrt(sum(1/input_matrix))
    output_odds2 <- exp(log(output_odds1)-output_std*1.96)
    output_odds3 <- exp(log(output_odds1)+output_std*1.96)  
    out_odds_ratio <- c(output_odds1,output_odds2,output_odds3)
    return(out_odds_ratio) 
  }
}
odds_ratio1 <- function(input_matrix){
  Output_matrix <- NULL
  for(mat_i in 2:nrow(input_matrix)){
    Output_matrix <- rbind(Output_matrix,
                           unit_odds_ratio(input_matrix[c(1,mat_i),]))
  }
  Output_matrix <- CI_trans(Output_matrix)
  Output_matrix <- rbind(c("---", "---"),
                         Output_matrix)
  return(Output_matrix)
}
# Calculate odds ratio with logistic regression
CI_trans <- function(input_matrix){
  input_A <- round(input_matrix[,1], 4)
  input_B <- paste0("[",
                    round(input_matrix[,2], 4),
                    ", ",
                    round(input_matrix[,3], 4),
                    "]")
  output_matrix <- cbind(input_A,input_B)
  return(output_matrix)
}

# Categorical Outcomes
table2_categorical <- function(Input_Data, ind_var_list, i, wb){
  # part 1: Basic summary statistics
  Output_data_1 <- NULL
  Output_data_2 <- NULL
  for (j in ind_var_list){
    # categorical predictors
    if(is.factor(Input_Data[,j])){
      table_unit <- table(Input_Data[,j],Input_Data[,i])
      # check assumption
      if(mean(table_unit<5)<=0.2){
        Output_unit <- c(i, 
                         j,
                         "Chi-squared test", 
                         round(chisq.test(table_unit)$p.value,4))
      }
      else{
        # high dimension table would be problematic, simulation would be used
        Fisher_results <- tryCatch(round(fisher.test(table_unit, simulate.p.value = TRUE)$p.value,4),
                                   error = function(e) return(NULL))
        if(!is.null(Fisher_results)){
          Output_unit <- c(i, 
                           j,
                           "Fisher's exact test (Simulation)",
                           Fisher_results)
        }
        else{
          Output_unit <- c(i, 
                           j,
                           "Fisher's exact test",
                           round(fisher.test(table_unit)$p.value,4))
        }
      }
      OR_table <- odds_ratio1(table_unit)
      Output_table <- cbind(row.names(table_unit), 
                            table_unit[,1],
                            round(prop.table(table_unit,2)[,1]*100, 2),
                            table_unit[,2],
                            round(prop.table(table_unit,2)[,2]*100, 2),
                            OR_table)

      Output_unit <- cbind(c(Output_unit[2], rep("",nrow(Output_table)-1)),
                           Output_table,
                           c(Output_unit[3], rep("",nrow(Output_table)-1)),
                           c(Output_unit[4], rep("",nrow(Output_table)-1)))
      Output_data_1 <- rbind(Output_data_1,Output_unit)
    }
    # continuous predictors
    else{
      # check assumption
      if(min(as.numeric(aggregate(Input_Data[,j],list(Input_Data[,i]),"shapiro.test")[,2][,2]))<0.05){
        Output_unit <- c("M-W test",
                         round(wilcox.test(Input_Data[,j]~Input_Data[,i])$p.value,4))
      }
      else{
        Output_unit <- c("t-test",
                         round(t.test(Input_Data[,j]~Input_Data[,i])$p.value,4))
      }
      Output_table <- cbind(round(aggregate(Input_Data[,j],list(Input_Data[,i]),"mean",na.rm=TRUE)[,2],2),
                            round(aggregate(Input_Data[,j],list(Input_Data[,i]),"sd",na.rm=TRUE)[,2],2),
                            round(aggregate(Input_Data[,j],list(Input_Data[,i]),"median",na.rm=TRUE)[,2],2),
                            paste0("[",
                                   round(aggregate(Input_Data[,j],list(Input_Data[,i]),"quantile",na.rm=TRUE)[,2][,2],2),
                                   ", ",
                                   round(aggregate(Input_Data[,j],list(Input_Data[,i]),"quantile",na.rm=TRUE)[,2][,4],2), 
                                   "]")
                            )
      Output_unit <- c(j,
                       Output_table[1,],
                       Output_table[2,],
                       Output_unit)
      Output_data_2 <- rbind(Output_data_2, Output_unit)
    }
  }
  if(!is.null(Output_data_1)){
    Output_data_1 <- data.frame(Output_data_1)
    colnames(Output_data_1) <- c("Independent Variable",
                                 "Levels",
                                 as.vector(t(outer(as.character(levels(Input_Data[, i])), 
                                                   c("Frequency", "Percent"), 
                                                   paste, sep=" "))),
                                 "ORs",
                                 "OR 95% CIs",
                                 "Test",
                                 "P-value")
  }
  if(!is.null(Output_data_2)){
    Output_data_2 <- data.frame(Output_data_2)
    colnames(Output_data_2) <- c("Independent Variable",
                                 as.vector(t(outer(as.character(levels(Input_Data[, i])), 
                                                   c("mean", "sd", "median", "IQR"), 
                                                   paste, sep=" "))),
                                 "Test",
                                 "P-value")
  }
  
  # part 2: logistic regression
  Output_data_3 <- NULL
  for(j in ind_var_list){
    if(is.factor(Input_Data[,j])){
      glm.model <- glm(Input_Data[,i]~as.factor(Input_Data[,j]),family=binomial(link="logit"))
      OR_matrix <- matrix(round(as.numeric(cbind(exp(cbind(coef(glm.model), 
                                                           confint(glm.model))))),
                                4),
                          ncol = 3, 
                          byrow = F)
      Coef_matrix <- round(summary(glm.model)$coef, 4)
      Output_unit <- cbind(Coef_matrix,
                           CI_trans(OR_matrix))
      ### exclude intercept
      Output_unit <- cbind(c("",j, rep("", length(levels(Input_Data[,j])[-1])-1)),
                           levels(Input_Data[,j]),
                           Output_unit)
      Output_unit <- Output_unit[-1,]
    }
    else{
      glm.model <- glm(Input_Data[,i]~Input_Data[,j],family=binomial(link="logit"))
      OR_matrix <- matrix(round(as.numeric(cbind(exp(cbind(coef(glm.model), 
                                                           confint(glm.model))))),
                                4),
                          ncol = 3, 
                          byrow = F)
      Coef_matrix <- round(summary(glm.model)$coef, 4)
      Output_unit <- cbind(Coef_matrix,
                       CI_trans(OR_matrix))
      ### exclude intercept
      Output_unit <- Output_unit[-1,]
      Output_unit <- c(j,
                       "---",
                       Output_unit)
    }
    Output_data_3 <- rbind(Output_data_3, 
                           Output_unit)
  }
  Output_data_3 <- data.frame(Output_data_3)
  colnames(Output_data_3) <- c("Independent Variable",
                               "Level",
                               "Coefficients",
                               "Std Error",
                               "Z-value",
                               "P-value",
                               "ORs",
                               "OR 95% CIs")
  
  
  # generate output
  addWorksheet(wb,i)
  Output_2_start <- 1
  Output_3_start <- 1
  if(!is.null(Output_data_1)){
    writeData(wb, i, 
              startCol = 1,
              startRow = 1, 
              x = "Categorical Variables")
    writeData(wb, i, 
              startCol = 1, 
              startRow = 2, 
              x = Output_data_1)
    Output_2_start <- nrow(Output_data_1)+5
  }
  if(!is.null(Output_data_2)){
    writeData(wb, i, 
              startCol = 1,
              startRow = Output_2_start, 
              x = "Continuous Variables")
    writeData(wb, i, 
              startCol = 1, 
              startRow = Output_2_start + 1, 
              x = Output_data_2)
    Output_3_start <- nrow(Output_data_2)+5
  }
  
  writeData(wb, i, 
            startCol = 1,
            startRow = Output_2_start + Output_3_start, 
            x = "Logistic regression models")
  writeData(wb, i, 
            startCol = 1, 
            startRow = Output_2_start + 1 + Output_3_start, 
            x = Output_data_3)
}

# Continuous Outcomes
table3_continuous <- function(Input_Data, ind_var_list, i, wb){
  # part 2: linear regression
  Output_data <- NULL
  for(j in ind_var_list){
    if(is.factor(Input_Data[,j])){
      lm.model <- tryCatch(lm(Input_Data[,i]~as.factor(Input_Data[,j])),
                           error = function(e){
                             return(NULL)
                           })
      if(is.null(lm.model)){
        Output_unit <- NULL
      }
      else{
        Coef_matrix <- round(summary(lm.model)$coef, 4)
        ### exclude intercept
        Output_unit <- cbind(c("",j, rep("", length(levels(Input_Data[,j])[-1])-1)),
                             levels(Input_Data[,j]),
                             Coef_matrix,
                             c("", 
                               round(summary(lm.model)$r.squared, 4), 
                               rep("", length(levels(Input_Data[,j])[-1])-1)),
                             c("",
                               round(summary(lm.model)$adj, 4),
                               rep("", length(levels(Input_Data[,j])[-1])-1)),
                             c("",
                               round(shapiro.test(lm.model$residuals)$p.value, 4), 
                               rep("", length(levels(Input_Data[,j])[-1])-1)))
        Output_unit <- Output_unit[-1,]
      }
    }
    else{
      lm.model <- lm(Input_Data[,i]~Input_Data[,j])
      Coef_matrix <- round(summary(lm.model)$coef, 4)
      ### exclude intercept
      Output_unit <- c(Coef_matrix[-1,], 
                       round(summary(lm.model)$r.squared, 4),
                       round(summary(lm.model)$adj, 4),
                       round(shapiro.test(lm.model$residuals)$p.value, 4))
      Output_unit <- c(j,
                       "---",
                       Output_unit)
    }
    Output_data <- rbind(Output_data, 
                           Output_unit)
  }
  Output_data <- data.frame(Output_data)
  colnames(Output_data) <- c("Independent Variable",
                               "Level",
                               "Coefficients",
                               "Std Error",
                               "T-value",
                               "P-value",
                               "R-squared",
                               "Adjusted r-squared",
                               "P-value for residual normality assumption")
  
  # generate output
  addWorksheet(wb,i)
  writeData(wb, i, 
            startCol = 1,
            startRow = 1, 
            x = "Linear Regression")
  writeData(wb, i, 
            startCol = 1, 
            startRow = 2, 
            x = Output_data)
  
  
  # provide plot
  for(j in 1:length(ind_var_list)){
    if(is.factor(Input_Data[,ind_var_list[j]])){
      lm.model <- tryCatch(lm(Input_Data[,i]~as.factor(Input_Data[,ind_var_list[j]])),
                           error = function(e){
                             return(NULL)
                           })
    }
    else{
      lm.model <- lm(Input_Data[,i]~Input_Data[,ind_var_list[j]])
    }
    if(!is.null(lm.model)){
      writeData(wb, i, 
                startCol = 1,
                startRow = nrow(Output_data)+6 + 22 * j - 22, 
                x = paste0("Assumption Check for predictor ", ind_var_list[j]))
      
      png(paste0("qqplot1_", j, "_", i, ".png"), width=768, height=768, units="px", res=144)
      Output_plot_unit <-  plot(lm.model, which = 1)
      dev.off()
      insertImage(wb, i, paste0("qqplot1_", j, "_", i, ".png"), width=4, height=4,
                  startCol = 1, 
                  startRow = nrow(Output_data)+ 7 + 22 * j - 22)
      
      png(paste0("qqplot2_", j, "_", i, ".png"), width=768, height=768, units="px", res=144)
      Output_plot_unit <-  plot(lm.model, which = 2)
      dev.off()
      insertImage(wb, i, paste0("qqplot2_", j, "_", i, ".png"), width=4, height=4,
                  startCol = 6, 
                  startRow = nrow(Output_data)+ 7 + 22 * j - 22)
      
      png(paste0("qqplot3_", j, "_", i, ".png"), width=768, height=768, units="px", res=144)
      Output_plot_unit <-  plot(lm.model, which = 3)
      dev.off()
      insertImage(wb, i, paste0("qqplot3_", j, "_", i, ".png"), width=4, height=4,
                  startCol = 11, 
                  startRow = nrow(Output_data)+ 7 + 22 * j - 22)
      
      png(paste0("qqplot4_", j, "_", i, ".png"), width=768, height=768, units="px", res=144)
      Output_plot_unit <-  plot(lm.model, which = 4)
      dev.off()
      insertImage(wb, i, paste0("qqplot4_", j, "_", i, ".png"), width=4, height=4,
                  startCol = 16, 
                  startRow = nrow(Output_data)+ 7 + 22 * j - 22)
      
    }
  }
}



#===========================================================================================================#

#=================================================    END       ============================================#

#===========================================================================================================#