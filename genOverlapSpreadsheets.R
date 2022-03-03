library(tidyverse)
library(edgeR)
library(statmod)
library(openxlsx)
library(VennDiagram)
library(gridExtra)
source("~/Grad School Docs/Research/scripts/degSpreadScripts/genDEG.R")

rs_rsComparsion <- function(degobj1 = "object generated from genDeg", 
                            degobj2 = "object generated from genDeg",
                            template = "xlsx template file of degoverlap",
                            filename = "name of output file"
                                 ){
  #make symbol of logFC to use in filter queries
  deg1.lfc <- paste("logFC", degobj1$`Contrast Condition`, sep = ".") %>% as.symbol()
  deg2.lfc <- paste("logFC", degobj2$`Contrast Condition`, sep = ".") %>% as.symbol()
  
  #union all present genes between two experiments
  unionFull <- full_join(degobj1$`All Present Genes`, 
                         degobj2$`All Present Genes`,
                         by = c("gene_id", "SYMBOL", "DESCRIPTION"),
                         suffix = paste(".",c(degobj1$`Contrast Condition`, degobj2$`Contrast Condition`), sep = ""))
  
  #intersect all present genes between two experiments
  interFull <- inner_join(degobj1$`All Present Genes`, 
                          degobj2$`All Present Genes`,
                          by = c("gene_id", "SYMBOL", "DESCRIPTION"),
                          suffix = paste(".",c(degobj1$`Contrast Condition`, degobj2$`Contrast Condition`), sep = ""))
  
  #union stat sig genes between two experiments
  unionSs <- full_join(degobj1$`Statistically Significant`, 
                       degobj2$`Statistically Significant`,
                       by = c("gene_id", "SYMBOL", "DESCRIPTION"),
                       suffix = paste(".",c(degobj1$`Contrast Condition`, degobj2$`Contrast Condition`), sep = ""))
  
  #intersect stat sig genes between two experiments
  interSs <- inner_join(degobj1$`Statistically Significant`, 
                        degobj2$`Statistically Significant`,
                        by = c("gene_id", "SYMBOL", "DESCRIPTION"),
                        suffix = paste(".",c(degobj1$`Contrast Condition`, degobj2$`Contrast Condition`), sep = ""))
  
  #list stat sig dataframes that contain concordant and discordant expression patterns
  subSsUnionList <- list(
    unionSs %>% filter(!!deg1.lfc > 0 & !!deg2.lfc > 0 ),
    unionSs %>% filter(!!deg1.lfc > 0 & !!deg2.lfc < 0 ),
    unionSs %>% filter(!!deg1.lfc < 0 & !!deg2.lfc > 0 ),
    unionSs %>% filter(!!deg1.lfc < 0 & !!deg2.lfc < 0 )
  )
  
  #name the dataframes
  names(subSsUnionList) <- c(sprintf("SS %s_Up %s_Up", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                             sprintf("SS %s_Up %s_Dn", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                             sprintf("SS %s_Dn %s_Up", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                             sprintf("SS %s_Dn %s_Dn", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`)
  )
  
  #list stat sig dataframes that contain expression patterns where experiment 1 observed expression but experiment 2 did not
  #and vice versa
  subSsWithNA <- list(
    unionSs %>% filter(!!deg1.lfc > 0 & is.na(!!deg2.lfc)),
    unionSs %>% filter(!!deg1.lfc < 0 & is.na(!!deg2.lfc)),
    unionSs %>% filter(is.na(!!deg1.lfc) & !!deg2.lfc > 0),
    unionSs %>% filter(is.na(!!deg1.lfc) & !!deg2.lfc < 0)
  )
  
  #name the dataframes
  names(subSsWithNA) <- c(sprintf("SS %s_Up %s_NoExp", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                          sprintf("SS %s_Dn %s_NoExp", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                          sprintf("SS %s_NoExp %s_Up", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                          sprintf("SS %s_NoExp %s_Dn", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`)
  )
  
  #union bio sig genes between two experiments
  unionBs <- full_join(degobj1$`Biologically Significant`, 
                       degobj2$`Biologically Significant`,
                       by = c("gene_id", "SYMBOL", "DESCRIPTION"),
                       suffix = paste(".",c(degobj1$`Contrast Condition`, degobj2$`Contrast Condition`), sep = ""))
  
  #intersect bio sig genes between two experiments
  interBs <- inner_join(degobj1$`Biologically Significant`, 
                        degobj2$`Biologically Significant`,
                        by = c("gene_id", "SYMBOL", "DESCRIPTION"),
                        suffix = paste(".",c(degobj1$`Contrast Condition`, degobj2$`Contrast Condition`), sep = ""))
  
  #list bio sig dataframes that contain concordant and discordant expression patterns
  subBsUnionList <- list(
    unionBs %>% filter(!!deg1.lfc > 0 & !!deg2.lfc > 0 ),
    unionBs %>% filter(!!deg1.lfc > 0 & !!deg2.lfc < 0 ),
    unionBs %>% filter(!!deg1.lfc < 0 & !!deg2.lfc > 0 ),
    unionBs %>% filter(!!deg1.lfc < 0 & !!deg2.lfc < 0 )
  )
  
  #name the dataframes
  names(subBsUnionList) <- c(sprintf("BS %s_Up %s_Up", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                             sprintf("BS %s_Up %s_Dn", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                             sprintf("BS %s_Dn %s_Up", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                             sprintf("BS %s_Dn %s_Dn", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`)
  )
  
  #list stat sig dataframes that contain expression patterns where experiment 1 observed expression but experiment 2 did not
  #and vice versa
  subBsWithNA <- list(
    unionBs %>% filter(!!deg1.lfc > 0 & is.na(!!deg2.lfc)),
    unionBs %>% filter(!!deg1.lfc < 0 & is.na(!!deg2.lfc)),
    unionBs %>% filter(is.na(!!deg1.lfc) & !!deg2.lfc > 0),
    unionBs %>% filter(is.na(!!deg1.lfc) & !!deg2.lfc < 0)
  )
  
  #name the dataframes
  names(subBsWithNA) <- c(sprintf("BS %s_Up %s_NoExp", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                          sprintf("BS %s_Dn %s_NoExp", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                          sprintf("BS %s_NoExp %s_Up", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                          sprintf("BS %s_NoExp %s_Dn", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`)
  )
  
  #statsig overlap summary matrix
  ssOverlapSummary <- matrix(data = c(Reduce(`+`, lapply(subSsWithNA[1:2], nrow)),
                                      nrow(subSsWithNA[[1]]),
                                      nrow(subSsWithNA[[2]]),
                                      Reduce(`+`, lapply(subSsUnionList, nrow)),
                                      nrow(subSsUnionList[[1]]),
                                      nrow(subSsUnionList[[4]]),
                                      Reduce(`+`, lapply(subSsWithNA[3:4], nrow)),
                                      nrow(subSsWithNA[[3]]),
                                      nrow(subSsWithNA[[4]])
  ),
  nrow = 3,
  dimnames = list(c("Total", "Up", "Down"),
                  c(degobj1$`Contrast Condition`, "Both", degobj2$`Contrast Condition`))
  )
  
  #biosig overlap summary matrix
  bsOverlapSummary <- matrix(data = c(Reduce(`+`, lapply(subBsWithNA[1:2], nrow)),
                                      nrow(subBsWithNA[[1]]),
                                      nrow(subBsWithNA[[2]]),
                                      Reduce(`+`, lapply(subBsUnionList, nrow)),
                                      nrow(subBsUnionList[[1]]),
                                      nrow(subBsUnionList[[4]]),
                                      Reduce(`+`, lapply(subBsWithNA[3:4], nrow)),
                                      nrow(subBsWithNA[[3]]),
                                      nrow(subBsWithNA[[4]])
  ),
  nrow = 3,
  dimnames = list(c("Total", "Up", "Down"),
                  c(degobj1$`Contrast Condition`, "Both", degobj2$`Contrast Condition`))
  )
  
  #statsig overlap contingency table
  ssOContingency <- matrix(data = c(nrow(unionFull) - sum(ssOverlapSummary[1,]),
                                    ssOverlapSummary[1,1],
                                    ssOverlapSummary[1,3],
                                    ssOverlapSummary[1,2]),
                           nrow = 2,
                           dimnames = list(c(paste("noIN",degobj1$`Contrast Condition`,sep = "."), 
                                             paste("IN",degobj1$`Contrast Condition`,sep = ".")),
                                           c(paste("noIN",degobj2$`Contrast Condition`,sep = "."), 
                                             paste("IN",degobj2$`Contrast Condition`,sep = ".")))
  )
  
  #biosig overlap contingency table
  bsOContingency <- matrix(data = c(nrow(unionFull) - sum(bsOverlapSummary[1,]),
                                    bsOverlapSummary[1,1],
                                    bsOverlapSummary[1,3],
                                    bsOverlapSummary[1,2]),
                           nrow = 2,
                           dimnames = list(c(paste("noIN",degobj1$`Contrast Condition`,sep = "."), 
                                             paste("IN",degobj1$`Contrast Condition`,sep = ".")),
                                           c(paste("noIN",degobj2$`Contrast Condition`,sep = "."), 
                                             paste("IN",degobj2$`Contrast Condition`,sep = ".")))
  )
  
  #fisher test on contigency tables (one-sided)
  ssOtest <- fisher.test(ssOContingency, alternative = "greater")
  bsOtest <- fisher.test(bsOContingency, alternative = "greater")

  #ss Inter summary
  ssInterSummary <- matrix(data = unlist(lapply(subSsUnionList,nrow)),
                           nrow = 2,
                           byrow = TRUE,
                           dimnames = list(c(paste(degobj1$`Contrast Condition`, "Up",sep = "_"), 
                                             paste(degobj1$`Contrast Condition`, "Down",sep = "_")),
                                           c(paste(degobj2$`Contrast Condition`, "Up",sep = "_"), 
                                             paste(degobj2$`Contrast Condition`, "Down",sep = "_"))
                           )
  )
  
  #bs Inter summary
  bsInterSummary <- matrix(data = unlist(lapply(subBsUnionList,nrow)),
                           nrow = 2,
                           byrow = TRUE,
                           dimnames = list(c(paste(degobj1$`Contrast Condition`, "Up",sep = "_"), 
                                             paste(degobj1$`Contrast Condition`, "Down",sep = "_")),
                                           c(paste(degobj2$`Contrast Condition`, "Up",sep = "_"), 
                                             paste(degobj2$`Contrast Condition`, "Down",sep = "_"))
                           )
  )
  
  #fisher test on Inter tables (two-sided)
  ssItest <- fisher.test(ssInterSummary)
  bsItest <- fisher.test(bsInterSummary)
  
  #condense all data frames into a list of data frames
  frames <- list("Full Union" = unionFull, "Full Intersection" = interFull,
                 "Stat Sig Union" = unionSs, "Stat Sig Intersection" = interSs,
                 "Bio Sig Union" = unionBs, "Bio Sig Intersection" = interBs)
  
  frames <- append(frames, subSsWithNA, after = 4)
  frames <- append(frames, subSsUnionList, after = 4)
  
  frames <- append(frames, subBsWithNA, after = 14)
  frames <- append(frames, subBsUnionList, after = 14)
  
  #generate workbook
  wb <- loadWorkbook(template)

  #write names of data frames to Worksheet Specification section
  wb %>% writeData(sheet = "Data Description", x = names(frames), startCol = 2, startRow = 14)
  
  #write name of Experiment 1 
  wb %>% writeData(sheet = "Data Description", x = degobj1$`Sample Title`, startCol = 2, startRow = 51)
  
  #write out summary data of Experiment 1
  #write out summary of experiment 1 SS data explaining Total genes, Up-reg genes, and Dn-reg genes
  wb %>% writeData(sheet = "Data Description", x = matrix(c(degobj1$`Statistically Significant` %>% nrow(),
                                                            degobj1$`Statistically Significant` %>% filter(Fold_Change > 0) %>% nrow,
                                                            degobj1$`Statistically Significant` %>% filter(Fold_Change < 0) %>% nrow),
                                                          nrow = 1),
                   colNames = FALSE,
                   startCol = 3, startRow = 53)
  
  #write out summary of experiment 1 BS data explaining Total genes, Up-reg genes, and Dn-reg genes
  wb %>% writeData(sheet = "Data Description", x = matrix(c(degobj1$`Biologically Significant` %>% nrow(), 
                                                            degobj1$`Biologically Significant` %>% filter(Fold_Change > 0) %>% nrow,
                                                            degobj1$`Biologically Significant` %>% filter(Fold_Change < 0) %>% nrow),
                                                          nrow = 1),
                   colNames = FALSE,
                   startCol = 3, startRow = 54)
  
  #write out summary of experiment 1 All Present Genes data explaining Total genes, Up-reg genes, and Dn-reg genes
  wb %>% writeData(sheet = "Data Description", x = matrix(c(degobj1$`All Present Genes` %>% nrow(), 
                                                            degobj1$`All Present Genes` %>% filter(Fold_Change > 0) %>% nrow,
                                                            degobj1$`All Present Genes` %>% filter(Fold_Change < 0) %>% nrow),
                                                          nrow = 1),
                   colNames = FALSE,
                   startCol = 3, startRow = 55)
  
  #write name of Experiment 2
  wb %>% writeData(sheet = "Data Description", x = degobj2$`Sample Title`, startCol = 2, startRow = 57)
  
  #write out summary data of Experiment 2
  #write out summary of experiment 2 SS data explaining Total genes, Up-reg genes, and Dn-reg genes
  wb %>% writeData(sheet = "Data Description", x = matrix(c(degobj2$`Statistically Significant` %>% nrow(), 
                                                            degobj2$`Statistically Significant` %>% filter(Fold_Change > 0) %>% nrow,
                                                            degobj2$`Statistically Significant` %>% filter(Fold_Change < 0) %>% nrow),
                                                          nrow = 1),
                   colNames = FALSE,
                   startCol = 3, startRow = 59)
  
  #write out summary of experiment 2 BS data explaining Total genes, Up-reg genes, and Dn-reg genes
  wb %>% writeData(sheet = "Data Description", x = matrix(c(degobj2$`Biologically Significant` %>% nrow(), 
                                                            degobj2$`Biologically Significant` %>% filter(Fold_Change > 0) %>% nrow,
                                                            degobj2$`Biologically Significant` %>% filter(Fold_Change < 0) %>% nrow),
                                                          nrow = 1),
                   colNames = FALSE,
                   startCol = 3, startRow = 60)
  
  #write out summary of experiment 2 All Present Genes data explaining Total genes, Up-reg genes, and Dn-reg genes
  wb %>% writeData(sheet = "Data Description", x = matrix(c(degobj2$`All Present Genes` %>% nrow(), 
                                                            degobj2$`All Present Genes` %>% filter(Fold_Change > 0) %>% nrow,
                                                            degobj2$`All Present Genes` %>% filter(Fold_Change < 0) %>% nrow),
                                                          nrow = 1),
                   colNames = FALSE,
                   startCol = 3, startRow = 61)
  
  #write out summary data comparing statistical sig overlap of the experiments
  #write out matrix containing data that describes overlap of Total genes, Up-reg genes, and Dn-reg genes between the experiments
  wb %>% writeData(sheet = "Data Description", x = ssOverlapSummary, startCol = 2, startRow = 64, colNames = TRUE, rowNames = TRUE)
  
  #write out amount of contra-regulated genes to SS Overlap summary
  wb %>% writeData(sheet = "Data Description", x = Reduce(`+`, lapply(subSsUnionList[2:3], nrow)), startCol = 4, startRow = 68)
  
  #write out contingency table that describes overlap of genes in either experiment: describes noIN.A, noIN.B, IN.A, IN.B
  wb %>% writeData(sheet = "Data Description", x = ssOContingency, startCol = 2, startRow = 71, colNames = TRUE, rowNames = TRUE)
  
  #write out p.value of one-sided fisher's exact test
  wb %>% writeData(sheet = "Data Description", x = ssOtest$p.value, startCol = 3, startRow = 74)
  
  #write out stat sig intersection matrix
  wb %>% writeData(sheet = "Data Description", x = ssInterSummary, startCol = 2, startRow = 77, colNames = TRUE, rowNames = TRUE)
  
  #write out p.value of two-sided fisher's exact test
  wb %>% writeData(sheet = "Data Description", x = ssItest$p.value, startCol = 3, startRow = 80)
  
  #write out summary data comparing biological sig overlap of the experiments
  #write out matrix containing data that describes overlap of Total genes, Up-reg genes, and Dn-reg genes between the experiments
  wb %>% writeData(sheet = "Data Description", x = bsOverlapSummary, startCol = 2, startRow = 83, colNames = TRUE, rowNames = TRUE)
  
  #write out amount of contra-regulated genes to BS Overlap summary
  wb %>% writeData(sheet = "Data Description", x = Reduce(`+`, lapply(subBsUnionList[2:3], nrow)), startCol = 4, startRow = 87)
  
  #write out contingency table that describes overlap of genes in either experiment: describes noIN.A, noIN.B, IN.A, IN.B
  wb %>% writeData(sheet = "Data Description", x = bsOContingency, startCol = 2, startRow = 90, colNames = TRUE, rowNames = TRUE)
  
  #write out p.value of one-sided fisher's exact test
  wb %>% writeData(sheet = "Data Description", x = bsOtest$p.value, startCol = 3, startRow = 93)
  
  #write out bio sig intersection matrix
  wb %>% writeData(sheet = "Data Description", x = bsInterSummary, startCol = 2, startRow = 96, colNames = TRUE, rowNames = TRUE)
  
  #write out p.value of two-sided fisher's exact test
  wb %>% writeData(sheet = "Data Description", x = bsItest$p.value, startCol = 3, startRow = 99)
  
  #write out column names in column description section
  wb %>% writeData(sheet = "Data Description", x = colnames(unionFull), startCol = 3, startRow = 103)
  
  #add data frames as sheets to the wb and write the data to the sheets
  walk(.x = names(frames), ~addWorksheet(wb, .x))
  walk2(.x = frames, .y = names(frames), ~writeData(wb, sheet = .y, x = .x))
  
  #style each worksheet
  headerStyle <- createStyle(textDecoration = "bold",
                             halign = "center")
  walk2(.x = names(frames), .y = frames, ~addStyle(wb, sheet = .x, style = headerStyle, rows = 1, cols = 1:ncol(.y)))
  walk2(.x = names(frames), .y = frames, ~setColWidths(wb, sheet = .x, cols = 4:ncol(.y), widths = "auto"))
  walk2(.x = names(frames), .y = frames, ~setColWidths(wb, sheet = .x, cols = 1:2, widths = "auto"))
  walk(.x = names(frames), ~setColWidths(wb, sheet = .x, cols = 3, widths = 40))
  
  saveWorkbook(wb, file = filename, overwrite = TRUE)
  
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
}

rs_maComparison <- function(rnaobj, maobj, template, filename){
  
  deg1.lfc <- paste("logFC", rnaobj$`Contrast Condition`, sep = ".") %>% as.symbol()
  deg2.lfc <- paste("logFC", maobj$`Contrast Condition`, sep = ".") %>% as.symbol()
  
  #full union
  unionFull <- full_join(rnaobj$`All Present Genes`, 
                         maobj$`Unique Values`,
                         by = c("gene_id", "SYMBOL", "DESCRIPTION"),
                         suffix = paste(".",c(rnaobj$`Contrast Condition`, maobj$`Contrast Condition`), sep = "")) %>% 
    dplyr::select(-ProbeID)
  
  interFull <- inner_join(rnaobj$`All Present Genes`, 
                          maobj$`Unique Values`,
                          by = c("gene_id", "SYMBOL", "DESCRIPTION"),
                          suffix = paste(".",c(rnaobj$`Contrast Condition`, maobj$`Contrast Condition`), sep = "")) %>% 
    dplyr::select(-ProbeID)
  
  #statistically significant
  unionSs <- full_join(rnaobj$`Statistically Significant`, 
                       maobj$`Statistically Significant`,
                       by = c("gene_id", "SYMBOL", "DESCRIPTION"),
                       suffix = paste(".",c(rnaobj$`Contrast Condition`, maobj$`Contrast Condition`), sep = "")) %>% 
    dplyr::select(-ProbeID)
  
  interSs <- inner_join(rnaobj$`Statistically Significant`, 
                        maobj$`Statistically Significant`,
                        by = c("gene_id", "SYMBOL", "DESCRIPTION"),
                        suffix = paste(".",c(rnaobj$`Contrast Condition`, maobj$`Contrast Condition`), sep = "")) %>% 
    dplyr::select(-ProbeID)
  
  subSsUnionList <- list(
    unionSs %>% filter(!!deg1.lfc > 0 & !!deg2.lfc > 0 ),
    unionSs %>% filter(!!deg1.lfc > 0 & !!deg2.lfc < 0 ),
    unionSs %>% filter(!!deg1.lfc < 0 & !!deg2.lfc > 0 ),
    unionSs %>% filter(!!deg1.lfc < 0 & !!deg2.lfc < 0 )
  )
  
  names(subSsUnionList) <- c(sprintf("SS %s_Up %s_Up", rnaobj$`Contrast Condition`, maobj$`Contrast Condition`),
                             sprintf("SS %s_Up %s_Dn", rnaobj$`Contrast Condition`, maobj$`Contrast Condition`),
                             sprintf("SS %s_Dn %s_Up", rnaobj$`Contrast Condition`, maobj$`Contrast Condition`),
                             sprintf("SS %s_Dn %s_Dn", rnaobj$`Contrast Condition`, maobj$`Contrast Condition`)
  )
  
  subSsWithNA <- list(
    unionSs %>% filter(!!deg1.lfc > 0 & is.na(!!deg2.lfc)),
    unionSs %>% filter(!!deg1.lfc < 0 & is.na(!!deg2.lfc)),
    unionSs %>% filter(is.na(!!deg1.lfc) & !!deg2.lfc > 0),
    unionSs %>% filter(is.na(!!deg1.lfc) & !!deg2.lfc < 0)
  )
  
  names(subSsWithNA) <- c(sprintf("SS %s_Up %s_NoExp", rnaobj$`Contrast Condition`, maobj$`Contrast Condition`),
                          sprintf("SS %s_Dn %s_NoExp", rnaobj$`Contrast Condition`, maobj$`Contrast Condition`),
                          sprintf("SS %s_NoExp %s_Up", rnaobj$`Contrast Condition`, maobj$`Contrast Condition`),
                          sprintf("SS %s_NoExp %s_Dn", rnaobj$`Contrast Condition`, maobj$`Contrast Condition`)
  )
  
  
  #biologically significant
  unionBs <- full_join(rnaobj$`Biologically Significant`, 
                       maobj$`Biologically Significant`,
                       by = c("gene_id", "SYMBOL", "DESCRIPTION"),
                       suffix = paste(".",c(rnaobj$`Contrast Condition`, maobj$`Contrast Condition`), sep = "")) %>% 
    dplyr::select(-ProbeID)
  
  interBs <- inner_join(rnaobj$`Biologically Significant`, 
                        maobj$`Biologically Significant`,
                        by = c("gene_id", "SYMBOL", "DESCRIPTION"),
                        suffix = paste(".",c(rnaobj$`Contrast Condition`, maobj$`Contrast Condition`), sep = "")) %>% 
    dplyr::select(-ProbeID)
  
  subBsUnionList <- list(
    unionBs %>% filter(!!deg1.lfc > 0 & !!deg2.lfc > 0 ),
    unionBs %>% filter(!!deg1.lfc > 0 & !!deg2.lfc < 0 ),
    unionBs %>% filter(!!deg1.lfc < 0 & !!deg2.lfc > 0 ),
    unionBs %>% filter(!!deg1.lfc < 0 & !!deg2.lfc < 0 )
  )
  
  names(subBsUnionList) <- c(sprintf("BS %s_Up %s_Up", rnaobj$`Contrast Condition`, maobj$`Contrast Condition`),
                             sprintf("BS %s_Up %s_Dn", rnaobj$`Contrast Condition`, maobj$`Contrast Condition`),
                             sprintf("BS %s_Dn %s_Up", rnaobj$`Contrast Condition`, maobj$`Contrast Condition`),
                             sprintf("BS %s_Dn %s_Dn", rnaobj$`Contrast Condition`, maobj$`Contrast Condition`)
  )
  
  subBsWithNA <- list(
    unionBs %>% filter(!!deg1.lfc > 0 & is.na(!!deg2.lfc)),
    unionBs %>% filter(!!deg1.lfc < 0 & is.na(!!deg2.lfc)),
    unionBs %>% filter(is.na(!!deg1.lfc) & !!deg2.lfc > 0),
    unionBs %>% filter(is.na(!!deg1.lfc) & !!deg2.lfc < 0)
  )
  
  names(subBsWithNA) <- c(sprintf("BS %s_Up %s_NoExp", rnaobj$`Contrast Condition`, maobj$`Contrast Condition`),
                          sprintf("BS %s_Dn %s_NoExp", rnaobj$`Contrast Condition`, maobj$`Contrast Condition`),
                          sprintf("BS %s_NoExp %s_Up", rnaobj$`Contrast Condition`, maobj$`Contrast Condition`),
                          sprintf("BS %s_NoExp %s_Dn", rnaobj$`Contrast Condition`, maobj$`Contrast Condition`)
  )
  
  #statsig overlap summary matrix
  
  ssOverlapSummary <- matrix(data = c(Reduce(`+`, lapply(subSsWithNA[1:2], nrow)),
                                      nrow(subSsWithNA[[1]]),
                                      nrow(subSsWithNA[[2]]),
                                      Reduce(`+`, lapply(subSsUnionList, nrow)),
                                      nrow(subSsUnionList[[1]]),
                                      nrow(subSsUnionList[[4]]),
                                      Reduce(`+`, lapply(subSsWithNA[3:4], nrow)),
                                      nrow(subSsWithNA[[3]]),
                                      nrow(subSsWithNA[[4]])
  ),
  nrow = 3,
  dimnames = list(c("Total", "Up", "Down"),
                  c(rnaobj$`Contrast Condition`, "Both", maobj$`Contrast Condition`))
  )
  
  
  #biosig overlap summary matrix
  bsOverlapSummary <- matrix(data = c(Reduce(`+`, lapply(subBsWithNA[1:2], nrow)),
                                      nrow(subBsWithNA[[1]]),
                                      nrow(subBsWithNA[[2]]),
                                      Reduce(`+`, lapply(subBsUnionList, nrow)),
                                      nrow(subBsUnionList[[1]]),
                                      nrow(subBsUnionList[[4]]),
                                      Reduce(`+`, lapply(subBsWithNA[3:4], nrow)),
                                      nrow(subBsWithNA[[3]]),
                                      nrow(subBsWithNA[[4]])
  ),
  nrow = 3,
  dimnames = list(c("Total", "Up", "Down"),
                  c(rnaobj$`Contrast Condition`, "Both", maobj$`Contrast Condition`))
  )
  
  #statsig overlap contingency table
  ssOContingency <- matrix(data = c(nrow(unionFull) - sum(ssOverlapSummary[1,]),
                                    ssOverlapSummary[1,1],
                                    ssOverlapSummary[1,3],
                                    ssOverlapSummary[1,2]),
                           nrow = 2,
                           dimnames = list(c(paste("noIN",rnaobj$`Contrast Condition`,sep = "."), 
                                             paste("IN",rnaobj$`Contrast Condition`,sep = ".")),
                                           c(paste("noIN",maobj$`Contrast Condition`,sep = "."), 
                                             paste("IN",maobj$`Contrast Condition`,sep = ".")))
  )
  
  #biosig overlap contingency table
  bsOContingency <- matrix(data = c(nrow(unionFull) - sum(bsOverlapSummary[1,]),
                                    bsOverlapSummary[1,1],
                                    bsOverlapSummary[1,3],
                                    bsOverlapSummary[1,2]),
                           nrow = 2,
                           dimnames = list(c(paste("noIN",rnaobj$`Contrast Condition`,sep = "."), 
                                             paste("IN",rnaobj$`Contrast Condition`,sep = ".")),
                                           c(paste("noIN",maobj$`Contrast Condition`,sep = "."), 
                                             paste("IN",maobj$`Contrast Condition`,sep = ".")))
  )
  
  #fisher test on contigency tables (one-sided)
  ssOtest <- fisher.test(ssOContingency, alternative = "greater")
  bsOtest <- fisher.test(bsOContingency, alternative = "greater")
  
  #ss Inter summary
  ssInterSummary <- matrix(data = unlist(lapply(subSsUnionList,nrow)),
                           nrow = 2,
                           byrow = TRUE,
                           dimnames = list(c(paste(rnaobj$`Contrast Condition`, "Up",sep = "_"), 
                                             paste(rnaobj$`Contrast Condition`, "Down",sep = "_")),
                                           c(paste(maobj$`Contrast Condition`, "Up",sep = "_"), 
                                             paste(maobj$`Contrast Condition`, "Down",sep = "_"))
                           )
  )
  
  #bs Inter summary
  bsInterSummary <- matrix(data = unlist(lapply(subBsUnionList,nrow)),
                           nrow = 2,
                           byrow = TRUE,
                           dimnames = list(c(paste(rnaobj$`Contrast Condition`, "Up",sep = "_"), 
                                             paste(rnaobj$`Contrast Condition`, "Down",sep = "_")),
                                           c(paste(maobj$`Contrast Condition`, "Up",sep = "_"), 
                                             paste(maobj$`Contrast Condition`, "Down",sep = "_"))
                           )
  )
  
  
  #fisher test on Inter tables (two-sided)
  ssItest <- fisher.test(ssInterSummary)
  bsItest <- fisher.test(bsInterSummary)
  
  #condense all data frames into a list of data frames
  frames <- list("Full Union" = unionFull, "Full Intersection" = interFull,
                 "Stat Sig Union" = unionSs, "Stat Sig Intersection" = interSs,
                 "Bio Sig Union" = unionBs, "Bio Sig Intersection" = interBs)
  
  frames <- append(frames, subSsWithNA, after = 4)
  frames <- append(frames, subSsUnionList, after = 4)
  
  frames <- append(frames, subBsWithNA, after = 14)
  frames <- append(frames, subBsUnionList, after = 14)
  
  #generate workbook
  wb <- loadWorkbook(template)
  
  #write names of data frames to Worksheet Specification section
  wb %>% writeData(sheet = "Data Description", x = names(frames), startCol = 2, startRow = 14)
  
  #write name of Experiment 1 
  wb %>% writeData(sheet = "Data Description", x = rnaobj$`Sample Title`, startCol = 2, startRow = 51)
  
  #write out summary data of Experiment 1
  #write out summary of SS data explaining Total genes, Up-reg genes, and Dn-reg genes
  wb %>% writeData(sheet = "Data Description", x = matrix(c(rnaobj$`Statistically Significant` %>% nrow(),
                                                            rnaobj$`Statistically Significant` %>% filter(Fold_Change > 0) %>% nrow,
                                                            rnaobj$`Statistically Significant` %>% filter(Fold_Change < 0) %>% nrow),
                                                          nrow = 1),
                   colNames = FALSE,
                   startCol = 3, startRow = 53)
  
  #write out summary of BS data explaining Total genes, Up-reg genes, and Dn-reg genes
  wb %>% writeData(sheet = "Data Description", x = matrix(c(rnaobj$`Biologically Significant` %>% nrow(), 
                                                            rnaobj$`Biologically Significant` %>% filter(Fold_Change > 0) %>% nrow,
                                                            rnaobj$`Biologically Significant` %>% filter(Fold_Change < 0) %>% nrow),
                                                          nrow = 1),
                   colNames = FALSE,
                   startCol = 3, startRow = 54)
  
  #write out summary of All Present Genes data explaining Total genes, Up-reg genes, and Dn-reg genes
  wb %>% writeData(sheet = "Data Description", x = matrix(c(rnaobj$`All Present Genes` %>% nrow(), 
                                                            rnaobj$`All Present Genes` %>% filter(Fold_Change > 0) %>% nrow,
                                                            rnaobj$`All Present Genes` %>% filter(Fold_Change < 0) %>% nrow),
                                                          nrow = 1),
                   colNames = FALSE,
                   startCol = 3, startRow = 55)
  
  #write name of Experiment 2
  wb %>% writeData(sheet = "Data Description", x = maobj$`Sample Title`, startCol = 2, startRow = 57)
  
  #write out summary data of Experiment 2
  #write out summary of SS data explaining Total genes, Up-reg genes, and Dn-reg genes
  wb %>% writeData(sheet = "Data Description", x = matrix(c(maobj$`Statistically Significant` %>% nrow(), 
                                                            maobj$`Statistically Significant` %>% filter(Fold_Change > 0) %>% nrow,
                                                            maobj$`Statistically Significant` %>% filter(Fold_Change < 0) %>% nrow),
                                                          nrow = 1),
                   colNames = FALSE,
                   startCol = 3, startRow = 59)
  
  #write out summary of BS data explaining Total genes, Up-reg genes, and Dn-reg genes
  wb %>% writeData(sheet = "Data Description", x = matrix(c(maobj$`Biologically Significant` %>% nrow(), 
                                                            maobj$`Biologically Significant` %>% filter(Fold_Change > 0) %>% nrow,
                                                            maobj$`Biologically Significant` %>% filter(Fold_Change < 0) %>% nrow),
                                                          nrow = 1),
                   colNames = FALSE,
                   startCol = 3, startRow = 60)
  
  #write out summary of Unique Values data explaining Total genes, Up-reg genes, and Dn-reg genes
  wb %>% writeData(sheet = "Data Description", x = matrix(c(maobj$`Unique Values` %>% nrow(), 
                                                            maobj$`Unique Values` %>% filter(Fold_Change > 0) %>% nrow,
                                                            maobj$`Unique Values` %>% filter(Fold_Change < 0) %>% nrow),
                                                          nrow = 1),
                   colNames = FALSE,
                   startCol = 3, startRow = 61)
  
  
  #write out summary data comparing statistical sig overlap of the experiments
  #write out matrix containing data that describes overlap of Total genes, Up-reg genes, and Dn-reg genes between the experiments
  wb %>% writeData(sheet = "Data Description", x = ssOverlapSummary, startCol = 2, startRow = 64, colNames = TRUE, rowNames = TRUE)
  
  #write out amount of contra-regulated genes to SS Overlap summary
  wb %>% writeData(sheet = "Data Description", x = Reduce(`+`, lapply(subSsUnionList[2:3], nrow)), startCol = 4, startRow = 68)
  
  #write out contingency table that describes overlap of genes in either experiment: describes noIN.A, noIN.B, IN.A, IN.B
  wb %>% writeData(sheet = "Data Description", x = ssOContingency, startCol = 2, startRow = 71, colNames = TRUE, rowNames = TRUE)
  
  #write out p.value of one-sided fisher's exact test
  wb %>% writeData(sheet = "Data Description", x = ssOtest$p.value, startCol = 3, startRow = 74)
  
  #write out summary data comparing biological sig overlap of the experiments
  #write out matrix containing data that describes overlap of Total genes, Up-reg genes, and Dn-reg genes between the experiments
  wb %>% writeData(sheet = "Data Description", x = bsOverlapSummary, startCol = 2, startRow = 77, colNames = TRUE, rowNames = TRUE)
  
  #write out amount of contra-regulated genes to BS Overlap summary
  wb %>% writeData(sheet = "Data Description", x = Reduce(`+`, lapply(subBsUnionList[2:3], nrow)), startCol = 4, startRow = 81)
  
  #write out contingency table that describes overlap of genes in either experiment: describes noIN.A, noIN.B, IN.A, IN.B
  wb %>% writeData(sheet = "Data Description", x = bsOContingency, startCol = 2, startRow = 84, colNames = TRUE, rowNames = TRUE)
  
  #write out p.value of one-sided fisher's exact test
  wb %>% writeData(sheet = "Data Description", x = bsOtest$p.value, startCol = 3, startRow = 87)
  
  #write out stat sig intersection matrix
  wb %>% writeData(sheet = "Data Description", x = ssInterSummary, startCol = 2, startRow = 90, colNames = TRUE, rowNames = TRUE)
  
  #write out p.value of two-sided fisher's exact test
  wb %>% writeData(sheet = "Data Description", x = ssItest$p.value, startCol = 3, startRow = 93)
  
  #write out bio sig intersection matrix
  wb %>% writeData(sheet = "Data Description", x = bsInterSummary, startCol = 2, startRow = 96, colNames = TRUE, rowNames = TRUE)
  
  #write out p.value of two-sided fisher's exact test
  wb %>% writeData(sheet = "Data Description", x = bsItest$p.value, startCol = 3, startRow = 99)
  
  #write out column names in column description section
  wb %>% writeData(sheet = "Data Description", x = colnames(unionFull), startCol = 3, startRow = 103)
  
  #add data frames as sheets to the wb and write the data to the sheets
  walk(.x = names(frames), ~addWorksheet(wb, .x))
  walk2(.x = frames, .y = names(frames), ~writeData(wb, sheet = .y, x = .x))
  
  #style each worksheet
  headerStyle <- createStyle(textDecoration = "bold",
                             halign = "center")
  walk2(.x = names(frames), .y = frames, ~addStyle(wb, sheet = .x, style = headerStyle, rows = 1, cols = 1:ncol(.y)))
  walk2(.x = names(frames), .y = frames, ~setColWidths(wb, sheet = .x, cols = 4:ncol(.y), widths = "auto"))
  walk2(.x = names(frames), .y = frames, ~setColWidths(wb, sheet = .x, cols = 1:2, widths = "auto"))
  walk(.x = names(frames), ~setColWidths(wb, sheet = .x, cols = 3, widths = 40))
  
  saveWorkbook(wb, file = filename, overwrite = TRUE)
  
}



