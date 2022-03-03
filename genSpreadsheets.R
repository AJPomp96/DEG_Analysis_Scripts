library(tidyverse)
library(edgeR)
library(statmod)
library(openxlsx)
library(VennDiagram)
library(gridExtra)
source("~/Grad School Docs/Research/scripts/degSpreadScripts/genDEG.R")

genrnaDEGSpread <- function(degobj, template, filename){
  wb <- loadWorkbook(template)
  headerStyle <- createStyle(textDecoration = "bold",
                             halign = "center")
  
  sheet_names <- c("All Present Genes", "Statistically Significant", "Biologically Significant")
  
  writeData(wb, sheet = "Data Description", startCol = 2, startRow = 23, x = degobj$`Sample Title`)
  
  writeData(wb, sheet = "Data Description", startCol = 3, startRow = 25, x = nrow(degobj$`Statistically Significant`))
  writeData(wb, sheet = "Data Description", startCol = 4, startRow = 25,
            x = degobj$`Statistically Significant` %>% filter(Fold_Change > 0) %>% nrow)
  writeData(wb, sheet = "Data Description", startCol = 5, startRow = 25,
            x = degobj$`Statistically Significant` %>% filter(Fold_Change < 0) %>% nrow)
  
  writeData(wb, sheet = "Data Description", startCol = 3, startRow = 26, x = nrow(degobj$`Biologically Significant`))
  writeData(wb, sheet = "Data Description", startCol = 4, startRow = 26,
            x = degobj$`Biologically Significant` %>% filter(Fold_Change > 0) %>% nrow)
  writeData(wb, sheet = "Data Description", startCol = 5, startRow = 26,
            x = degobj$`Biologically Significant` %>% filter(Fold_Change < 0) %>% nrow)
  
  addWorksheet(wb, sheet = "Full Count Matrix")
  writeData(wb, sheet = "Full Count Matrix", x = degobj$`Full Count Matrix`, colNames = TRUE)
  addStyle(wb, sheet = "Full Count Matrix", style =  headerStyle, rows = 1, cols = 1:ncol(degobj$`Full Count Matrix`))
  
  walk(sheet_names, ~addWorksheet(wb, .x))
  walk(sheet_names, ~writeData(wb, x = degobj[[.x]], sheet = .x, colNames = TRUE))
  walk(sheet_names, ~setColWidths(wb, sheet = .x, cols = 5:ncol(degobj[[.x]]), widths = "auto"))
  walk(.x = sheet_names, ~addStyle(wb, sheet = .x, style = headerStyle, rows = 1, cols = 1:ncol(degobj[[.x]])))
  saveWorkbook(wb, file = filename, overwrite = TRUE)
}

genmaDEGSpread <- function(degobj, template, filename){
  wb <- loadWorkbook(template)
  headerStyle <- createStyle(textDecoration = "bold",
                             halign = "center")
  
  sheet_names <- c("Unique Values", "Statistically Significant", "Biologically Significant")
  
  writeData(wb, sheet = "Data Description", startCol = 2, startRow = 23, x = degobj$`Sample Title`)
  
  writeData(wb, sheet = "Data Description", startCol = 3, startRow = 25, x = nrow(degobj$`Statistically Significant`))
  writeData(wb, sheet = "Data Description", startCol = 4, startRow = 25,
            x = degobj$`Statistically Significant` %>% filter(Fold_Change > 0) %>% nrow)
  writeData(wb, sheet = "Data Description", startCol = 5, startRow = 25,
            x = degobj$`Statistically Significant` %>% filter(Fold_Change < 0) %>% nrow)
  
  writeData(wb, sheet = "Data Description", startCol = 3, startRow = 26, x = nrow(degobj$`Biologically Significant`))
  writeData(wb, sheet = "Data Description", startCol = 4, startRow = 26,
            x = degobj$`Biologically Significant` %>% filter(Fold_Change > 0) %>% nrow)
  writeData(wb, sheet = "Data Description", startCol = 5, startRow = 26,
            x = degobj$`Biologically Significant` %>% filter(Fold_Change < 0) %>% nrow)
  
  addWorksheet(wb, sheet = "All Present Genes")
  writeData(wb, sheet = "All Present Genes", x = degobj$`All Present Genes`, colNames = TRUE)
  addStyle(wb, sheet = "All Present Genes", style =  headerStyle, rows = 1, cols = 1:ncol(degobj$`All Present Genes`))
  
  walk(sheet_names, ~addWorksheet(wb, .x))
  walk(sheet_names, ~writeData(wb, x = degobj[[.x]], sheet = .x, colNames = TRUE))
  walk(sheet_names, ~setColWidths(wb, sheet = .x, cols = 5:ncol(degobj[[.x]]), widths = "auto"))
  walk(.x = sheet_names, ~addStyle(wb, sheet = .x, style = headerStyle, rows = 1, cols = 1:ncol(degobj[[.x]])))
  saveWorkbook(wb, file = filename, overwrite = TRUE)
}