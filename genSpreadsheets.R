library(tidyverse)
library(edgeR)
library(statmod)
library(openxlsx)
library(VennDiagram)
library(gridExtra)
source("~/Grad School Docs/Research/scripts/degSpreadScripts/genDEG.R")

addOverlapSheets <- function(wb = "#workbook object", 
                      degOverlapObj = "list of dataframes generated from genDegOverlap or genDEG"
                      ){
  i=1
  for(item in degOverlapObj){
    if(class(item) == "data.frame"){
      headerStyle <- createStyle(textDecoration = "bold",
                                 halign = "center")
      
      wb %>% 
        addWorksheet(sheetName = names(degOverlapObj)[i])
      
      wb %>%
        writeData(x = item,
                  sheet = names(degOverlapObj)[i])
      
      wb %>% 
        addStyle(sheet = names(degOverlapObj)[i], 
                 headerStyle, 
                 cols = 1:ncol(item), 
                 rows = 1)
      
      wb %>%
        setColWidths(sheet = names(degOverlapObj)[i], 
                     widths = "auto",
                     cols = 1:ncol(item))
      
      i <- i + 1
    }
    if(class(item) == "list"){
      j=1
      for(sub in item){
        headerStyle <- createStyle(textDecoration = "bold",
                                   halign = "center")
        
        wb %>%
          addWorksheet(sheetName = names(item)[j])
        
        wb %>%
          writeData(x = sub,
                    sheet = names(item)[j])
        
        wb %>% 
          addStyle(sheet = names(item)[j], 
                   headerStyle, 
                   cols = 1:ncol(sub), 
                   rows = 1)
        
        wb %>%
          setColWidths(sheet = names(item)[j], 
                       widths = "auto",
                       cols = 1:ncol(sub))
        
        j <- j + 1
      }
      i <- i + 1
    }
  }
}


addDegOverlapSummarySheet <- function(degobj1 = "object generated from genDeg", 
                                 degobj2 = "object generated from genDeg",
                                 degOverlapObj = "overlapObj generated from genDegOverlap",
                                 templatefile = "xlsx template file of degoverlap"
                                 ){
  
  obj1Summary <- matrix(data = c(nrow(degobj1$`Statistically Significant`),
                                 nrow(degobj1$`Statistically Significant` %>% filter(logFC > 0)),
                                 nrow(degobj1$`Statistically Significant` %>% filter(logFC < 0)),
                                 nrow(degobj1$`Biologically Significant`),
                                 nrow(degobj1$`Biologically Significant` %>% filter(logFC > 0)),
                                 nrow(degobj1$`Biologically Significant` %>% filter(logFC < 0))    
                                      ),
                        nrow = 2,
                        byrow = TRUE,
                        dimnames = list(c("Statistically Significant","Biologically Significant"),
                                        c("Total", "Up", "Down"))
                        )
  
  obj2Summary <- matrix(data = c(nrow(degobj2$`Statistically Significant`),
                                 nrow(degobj2$`Statistically Significant` %>% filter(logFC > 0)),
                                 nrow(degobj2$`Statistically Significant` %>% filter(logFC < 0)),
                                 nrow(degobj2$`Biologically Significant`),
                                 nrow(degobj2$`Biologically Significant` %>% filter(logFC > 0)),
                                 nrow(degobj2$`Biologically Significant` %>% filter(logFC < 0))
                                 ),
                        nrow = 2,
                        byrow = TRUE,
                        dimnames = list(c("Statistically Significant","Biologically Significant"),
                                        c("Total", "Up", "Down"))
                        )
  
  ssOverlapSummary <- matrix(data = c(Reduce(`+`, lapply(degOverlapObj$`Subset SS w/ NA`[1:2], nrow)),
                                      nrow(degOverlapObj$`Subset SS w/ NA`[[1]]),
                                      nrow(degOverlapObj$`Subset SS w/ NA`[[2]]),
                                      Reduce(`+`, lapply(degOverlapObj$`Subset SS`[c(1,4)], nrow)),
                                      nrow(degOverlapObj$`Subset SS`[[1]]),
                                      nrow(degOverlapObj$`Subset SS`[[4]]),
                                      Reduce(`+`, lapply(degOverlapObj$`Subset SS w/ NA`[3:4], nrow)),
                                      nrow(degOverlapObj$`Subset SS w/ NA`[[3]]),
                                      nrow(degOverlapObj$`Subset SS w/ NA`[[4]])
                                      ),
                             nrow = 3,
                             dimnames = list(c("Total", "Up", "Down"),
                                             c(degobj1$`Contrast Condition`, "Both", degobj2$`Contrast Condition`))
                             )
  bsOverlapSummary <- matrix(data = c(Reduce(`+`, lapply(degOverlapObj$`Subset BS w/ NA`[1:2], nrow)),
                                      nrow(degOverlapObj$`Subset BS w/ NA`[[1]]),
                                      nrow(degOverlapObj$`Subset BS w/ NA`[[2]]),
                                      Reduce(`+`, lapply(degOverlapObj$`Subset BS`[c(1,4)], nrow)),
                                      nrow(degOverlapObj$`Subset BS`[[1]]),
                                      nrow(degOverlapObj$`Subset BS`[[4]]),
                                      Reduce(`+`, lapply(degOverlapObj$`Subset BS w/ NA`[3:4], nrow)),
                                      nrow(degOverlapObj$`Subset BS w/ NA`[[3]]),
                                      nrow(degOverlapObj$`Subset BS w/ NA`[[4]])
                                      ),
                             nrow = 3,
                             dimnames = list(c("Total", "Up", "Down"),
                                             c(degobj1$`Contrast Condition`, "Both", degobj2$`Contrast Condition`))
                             )
  
  ssInterSummary <- matrix(data = lapply(degOverlapObj$`Subset SS`,nrow),
                           nrow = 2,
                           byrow = TRUE,
                           dimnames = list(c(paste(degobj1$`Contrast Condition`,"_Up",sep = ""), 
                                             paste(degobj1$`Contrast Condition`,"_Down",sep = "")),
                                           c(paste(degobj2$`Contrast Condition`,"_Up",sep = ""), 
                                             paste(degobj2$`Contrast Condition`,"_Down",sep = ""))
                                           )
                           )
  
  bsInterSummary <- matrix(data = lapply(degOverlapObj$`Subset BS`,nrow),
                           nrow = 2,
                           byrow = TRUE,
                           dimnames = list(c(paste(degobj1$`Contrast Condition`,"_Up",sep = ""),
                                             paste(degobj1$`Contrast Condition`,"_Down",sep = "")),
                                           c(paste(degobj2$`Contrast Condition`,"_Up",sep = ""),
                                             paste(degobj2$`Contrast Condition`,"_Down",sep = ""))
                                           )
                           )
  
  wsDescription <- c(sprintf("Statistically significant genes, Upregulated in %s(%s) results, Upregulated in %s(%s) results",
                             degobj1$`Contrast Condition`,degobj1$`Sample Comparison`,
                             degobj2$`Contrast Condition`, degobj2$`Sample Comparison`),
                     sprintf("Statistically significant genes, Upregulated in %s(%s) results, Downregulated in %s(%s) results",
                             degobj1$`Contrast Condition`,degobj1$`Sample Comparison`,
                             degobj2$`Contrast Condition`, degobj2$`Sample Comparison`),
                     sprintf("Statistically significant genes, Downregulated in %s(%s) results, Upregulated in %s(%s) results",
                             degobj1$`Contrast Condition`,degobj1$`Sample Comparison`,
                             degobj2$`Contrast Condition`, degobj2$`Sample Comparison`),
                     sprintf("Statistically significant genes, Downregulated in %s(%s) results, Downregulated in %s(%s) results",
                             degobj1$`Contrast Condition`,degobj1$`Sample Comparison`,
                             degobj2$`Contrast Condition`, degobj2$`Sample Comparison`),
                     sprintf("Statistically significant genes, Upregulated in %s(%s) results with no expression observed in %s(%s) results",
                             degobj1$`Contrast Condition`,degobj1$`Sample Comparison`,
                             degobj2$`Contrast Condition`, degobj2$`Sample Comparison`),
                     sprintf("Statistically significant genes, Downregulated in %s(%s) results, with no expression observed in %s(%s) results",
                             degobj1$`Contrast Condition`,degobj1$`Sample Comparison`,
                             degobj2$`Contrast Condition`, degobj2$`Sample Comparison`),
                     sprintf("Statistically significant genes, Upregulated in %s(%s) results with no expression observed in %s(%s) results",
                             degobj2$`Contrast Condition`, degobj2$`Sample Comparison`,
                             degobj1$`Contrast Condition`,degobj1$`Sample Comparison`),
                     sprintf("Statistically significant genes, Downregulated in %s(%s) results, with no expression observed in %s(%s) results",
                             degobj2$`Contrast Condition`, degobj2$`Sample Comparison`,
                             degobj1$`Contrast Condition`,degobj1$`Sample Comparison`),
                     sprintf("Biologically significant genes, Upregulated in %s(%s) results, Upregulated in %s(%s) results",
                             degobj1$`Contrast Condition`,degobj1$`Sample Comparison`,
                             degobj2$`Contrast Condition`, degobj2$`Sample Comparison`),
                     sprintf("Biologically significant genes, Upregulated in %s(%s) results, Downregulated in %s(%s) results",
                             degobj1$`Contrast Condition`,degobj1$`Sample Comparison`,
                             degobj2$`Contrast Condition`, degobj2$`Sample Comparison`),
                     sprintf("Biologically significant genes, Downregulated in %s(%s) results, Upregulated in %s(%s) results",
                             degobj1$`Contrast Condition`,degobj1$`Sample Comparison`,
                             degobj2$`Contrast Condition`, degobj2$`Sample Comparison`),
                     sprintf("Biologically significant genes, Downregulated in %s(%s) results, Downregulated in %s(%s) results",
                             degobj1$`Contrast Condition`,degobj1$`Sample Comparison`,
                             degobj2$`Contrast Condition`, degobj2$`Sample Comparison`),
                     sprintf("Biologically significant genes, Upregulated in %s(%s) results, no expression observed in %s(%s) results",
                             degobj1$`Contrast Condition`,degobj1$`Sample Comparison`,
                             degobj2$`Contrast Condition`, degobj2$`Sample Comparison`),
                     sprintf("Biologically significant genes, Downregulated in %s(%s) results, no expression observed in %s(%s) results",
                             degobj1$`Contrast Condition`,degobj1$`Sample Comparison`,
                             degobj2$`Contrast Condition`, degobj2$`Sample Comparison`),
                     sprintf("Biologically significant genes, Upregulated in %s(%s) results, no expression observed in %s(%s) results",
                             degobj2$`Contrast Condition`, degobj2$`Sample Comparison`,
                             degobj1$`Contrast Condition`,degobj1$`Sample Comparison`),
                     sprintf("Biologically significant genes, Downregulated in %s(%s) results, no expression observed in %s(%s) results",
                             degobj2$`Contrast Condition`, degobj2$`Sample Comparison`,
                             degobj1$`Contrast Condition`,degobj1$`Sample Comparison`)
                     )
  
  colSummary <- c(sprintf("The fold change of this gene between  %s  and %s  observed in %s(%s)",
                          degobj1$`Condition 1`, degobj1$`Condition 2`, degobj1$`Contrast Condition`, degobj1$`Sample Comparison`),
                  sprintf("The log fold change of this gene between  %s  and %s  observed in %s(%s)",
                          degobj1$`Condition 1`, degobj1$`Condition 2`, degobj1$`Contrast Condition`, degobj1$`Sample Comparison`),
                  sprintf("P_Value calculated by edgeR %s(%s)",
                          degobj1$`Contrast Condition`, degobj1$`Sample Comparison`),
                  sprintf("Adjusted P value calculated by edgeR %s(%s)",
                          degobj1$`Contrast Condition`, degobj1$`Sample Comparison`),
                  sprintf("Abundance of this gene (Average FPKM) observed in %s, calculated by %s(%s)",
                          degobj1$`Condition 1`, degobj1$`Contrast Condition`, degobj1$`Sample Comparison`),
                  sprintf("Abundance of this gene (Average FPKM) observed in %s, calculated by %s(%s)",
                          degobj1$`Condition 2`, degobj1$`Contrast Condition`, degobj1$`Sample Comparison`),
                  sprintf("The fold change of this gene between %s and %s observed in %s(%s)",
                          degobj2$`Condition 1`, degobj2$`Condition 2`, degobj2$`Contrast Condition`, degobj2$`Sample Comparison`),
                  sprintf("The log fold change of this gene between  %s  and %s  observed in %s(%s)",
                          degobj2$`Condition 1`, degobj2$`Condition 2`, degobj2$`Contrast Condition`, degobj2$`Sample Comparison`),
                  sprintf("P_Value calculated by edgeR %s(%s)",
                          degobj2$`Contrast Condition`, degobj2$`Sample Comparison`),
                  sprintf("Adjusted P value calculated by edgeR %s(%s)",
                          degobj2$`Contrast Condition`, degobj2$`Sample Comparison`),
                  sprintf("Abundance of this gene (Average FPKM) observed in %s, calculated by %s(%s)",
                          degobj2$`Condition 1`, degobj2$`Contrast Condition`, degobj2$`Sample Comparison`),
                  sprintf("Abundance of this gene (Average FPKM) observed in %s, calculated by %s(%s)",
                          degobj2$`Condition 2`, degobj2$`Contrast Condition`, degobj2$`Sample Comparison`)
                  )
  
  bsTotalVenn <- draw.pairwise.venn(area1 = sum(bsOverlapSummary[1,1:2]), 
                                    area2 = sum(bsOverlapSummary[1,2:3]), 
                                    cross.area = bsOverlapSummary[1,2],
                                    category = c(degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                                    fontfamily = c("sans"),
                                    cat.fontfamily = c("sans"),
                                    lty = rep("blank", 2),
                                    fill = c("light blue", "pink"), 
                                    alpha = rep(0.5, 2), 
                                    cat.pos = c(0, 0), 
                                    cat.dist = rep(0.025, 2),
                                    cat.cex = c(.8))
  
  bsUpVenn <- draw.pairwise.venn(area1 = sum(bsOverlapSummary[2,1:2]), 
                                 area2 = sum(bsOverlapSummary[2,2:3]), 
                                 cross.area = bsOverlapSummary[2,2],
                                 category = c(degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                                 fontfamily = c("sans"),
                                 cat.fontfamily = c("sans"),
                                 lty = rep("blank", 2),
                                 fill = c("light blue", "pink"), 
                                 alpha = rep(0.5, 2), 
                                 cat.pos = c(0, 0), 
                                 cat.dist = rep(0.025, 2),
                                 cat.cex = c(.8))
  
  bsDnVenn <- draw.pairwise.venn(area1 = sum(bsOverlapSummary[3,1:2]), 
                                 area2 = sum(bsOverlapSummary[3,2:3]), 
                                 cross.area = bsOverlapSummary[3,2],
                                 category = c(degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                                 fontfamily = c("sans"),
                                 cat.fontfamily = c("sans"),
                                 lty = rep("blank", 2),
                                 fill = c("light blue", "pink"), 
                                 alpha = rep(0.5, 2), 
                                 cat.pos = c(0, 0), 
                                 cat.dist = rep(0.025, 2),
                                 cat.cex = c(.8))
  
  ssTotalVenn <- draw.pairwise.venn(area1 = sum(ssOverlapSummary[1,1:2]), 
                                    area2 = sum(ssOverlapSummary[1,2:3]), 
                                    cross.area = ssOverlapSummary[1,2],
                                    category = c(degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                                    fontfamily = c("sans"),
                                    cat.fontfamily = c("sans"),
                                    lty = rep("blank", 2),
                                    fill = c("light blue", "pink"), 
                                    alpha = rep(0.5, 2), 
                                    cat.pos = c(0, 0), 
                                    cat.dist = rep(0.025, 2),
                                    cat.cex = c(.8))
  
  ssUpVenn <- draw.pairwise.venn(area1 = sum(ssOverlapSummary[2,1:2]), 
                                 area2 = sum(ssOverlapSummary[2,2:3]), 
                                 cross.area = ssOverlapSummary[2,2],
                                 category = c(degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                                 fontfamily = c("sans"),
                                 cat.fontfamily = c("sans"),
                                 lty = rep("blank", 2),
                                 fill = c("light blue", "pink"), 
                                 alpha = rep(0.5, 2), 
                                 cat.pos = c(0, 0), 
                                 cat.dist = rep(0.025, 2),
                                 cat.cex = c(.8))
  
  ssDnVenn <- draw.pairwise.venn(area1 = sum(ssOverlapSummary[3,1:2]), 
                                 area2 = sum(ssOverlapSummary[3,2:3]), 
                                 cross.area = ssOverlapSummary[3,2],
                                 category = c(degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                                 fontfamily = c("sans"),
                                 cat.fontfamily = c("sans"),
                                 lty = rep("blank", 2),
                                 fill = c("light blue", "pink"), 
                                 alpha = rep(0.5, 2), 
                                 cat.pos = c(0, 0), 
                                 cat.dist = rep(0.025, 2),
                                 cat.cex = c(.8))
  
  dev.off()
  
  wb <- loadWorkbook(templatefile)
  
  wb %>%
    writeData(sheet = "Data Description",
              x = c(names(overlapObj[1]),
                    names(overlapObj$`Subset SS`),
                    names(overlapObj$`Subset SS w/ NA`),
                    names(overlapObj[4]),
                    names(overlapObj$`Subset BS`),
                    names(overlapObj$`Subset BS w/ NA`),
                    names(overlapObj[7])
                    ),
              startRow = 14,
              startCol = 2)
  
  wb %>%
    writeData(sheet = "Data Description",
              x = wsDescription[1:8],
              startRow = 15,
              startCol = 3)
  
  wb %>%
    writeData(sheet = "Data Description",
              x = wsDescription[9:16],
              startRow = 24,
              startCol = 3)
  
  wb %>%
    writeData(sheet = "Data Description",
              x = degobj1$`Sample Title`,
              startRow = 46,
              startCol = 2)
  
  wb %>% 
    writeData(sheet = "Data Description",
                   x = obj1Summary,
                   startRow = 47,
                   startCol = 2,
                   rowNames = TRUE
                   )
  
  wb %>%
    writeData(sheet = "Data Description",
              x = degobj2$`Sample Title`,
              startRow = 52,
              startCol = 2)
  
  wb %>%
    writeData(sheet = "Data Description",
              x = obj2Summary,
              startRow = 53,
              startCol = 2,
              rowNames = TRUE
              )
  
  wb %>%
    writeData(sheet = "Data Description",
              x = ssOverlapSummary,
              startRow = 59,
              startCol = 2,
              rowNames = TRUE)
  
  wb %>%
    writeData(sheet = "Data Description",
              x = bsOverlapSummary,
              startRow = 66,
              startCol = 2,
              rowNames = TRUE)
  
  wb %>%
    writeData(sheet = "Data Description",
              x = ssInterSummary,
              startRow = 73,
              startCol = 2,
              rowNames = TRUE
              )
  
  wb %>%
    writeData(sheet = "Data Description",
              x = bsInterSummary,
              startRow = 79,
              startCol = 2,
              rowNames = TRUE
              )
  
  wb %>%
    writeData(sheet = "Data Description",
              x = names(degOverlapObj$`Full Intersection`[2:15]),
              startRow = 85,
              startCol = 3
              )
  
  wb %>%
    writeData(sheet = "Data Description",
              x = colSummary,
              startRow = 87,
              startCol = 5)
  
  grid.arrange(gTree(children = bsTotalVenn), top = "BioSig Overlap of All Genes")
  
  insertPlot(wb = wb, sheet = "Data Description",
             startRow = 101, startCol = 1)
  
  dev.off()
  
  grid.arrange(gTree(children = bsUpVenn), top = "BioSig Overlap of Upregulated Genes")
  
  insertPlot(wb = wb, sheet = "Data Description",
             startRow = 121, startCol = 1)
  
  dev.off()
  
  grid.arrange(gTree(children = bsDnVenn), top = "BioSig Overlap of Downregulated Genes")
  
  insertPlot(wb = wb, sheet = "Data Description",
             startRow = 141, startCol = 1)
  
  dev.off()
  
  grid.arrange(gTree(children = ssTotalVenn), top = "StatSig Overlap of All Genes")
  
  insertPlot(wb = wb, sheet = "Data Description",
             startRow = 101, startCol = 6)
  
  dev.off()
  
  grid.arrange(gTree(children = ssUpVenn), top = "StatSig Overlap of Upregulated Genes")
  
  insertPlot(wb = wb, sheet = "Data Description",
             startRow = 121, startCol = 6)
  
  dev.off()
  
  grid.arrange(gTree(children = ssDnVenn), top = "StatSig Overlap of Downregulated Genes")
  
  insertPlot(wb = wb, sheet = "Data Description",
             startRow = 141, startCol = 6)
  
  dev.off()

  return(wb)
}

genOverlapSpread <- function(degobj1 = "deg object generated from genDEG", 
                             degobj2 = "deg object generated from genDEG"
                             ){
  wb <- addDegOverlapSummarySheet()
  addSheets(wb, overlapObj)
  saveWorkbook(wb, file = "test.xlsx", overwrite = TRUE)
}



