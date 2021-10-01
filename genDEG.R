library(tidyverse)
library(edgeR)
library(statmod)
library(openxlsx)

genSampleTable <- function(wd,
                           condition1,
                           condition2,
                           c1pattern
                           ){
  countFiles <- list.files(wd, pattern = "GeneCount")
  samples <- gsub("_rf_GeneCount.txt", "", countFiles)
  condition <- ifelse(grepl(c1pattern, samples), condition1, condition2)
  sampleTable <- data.frame(sampleName=samples, 
                            fileName=paste(wd,countFiles,sep = "/"), 
                            condition=condition)
  return(sampleTable)
}

genDEG <- function(degTitle = "title for the deg object. ex: WT vs cKO epithelium",
                   sampleComparison = "conditions compared. ex: WT vs cKO",
                   contrastCondition = "contrast condtion. only used if you are comparing two deg objs",
                   condition1 = "condition of the control replicates. ex: WT/control",
                   condition2 = "condition of the experimental replicates. ex: cKO/mutant",
                   sampleTable = "sample table generated using genSampleTable", 
                   geneAnnotPath = "path to gene Annotation file"
                   ){
  
  geneAnnot <- read.csv(geneAnnotPath, row.names = 1)
  row.names(geneAnnot) <- geneAnnot$gene_id
  
  dge <- readDGE(files = sampleTable$fileName,
                 group = sampleTable$condition,
                 labels = sampleTable$sampleName,
                 header = FALSE)
  
  dge$genes <- geneAnnot[row.names(dge$counts),]
  dge <- dge[!grepl("__",row.names(dge$counts)), ,keep.lib.sizes=FALSE]
  dge <- dge[filterByExpr(dge), ,keep.lib.sizes=FALSE]
  dge <- calcNormFactors(dge, method = "TMM")
  dge <- estimateDisp(dge, robust = TRUE)
  
  rbg <- as.data.frame(rpkmByGroup(dge, gene.length = "eu_length"))
  rbg$gene_id <- row.names(rbg)
  dge$genes <- merge(rbg, dge$genes, "gene_id")
  colnames(dge$genes)[colnames(dge$genes) == condition1] <- sprintf("%s_Avg_FPKM", condition1)
  colnames(dge$genes)[colnames(dge$genes) == condition2] <- sprintf("%s_Avg_FPKM", condition2)
  
  et <- exactTest(dge, pair = c(condition1,condition2))
  
  diffExpGenes <- as.data.frame(topTags(et, n = Inf))
  diffExpGenes <- diffExpGenes %>%
    mutate(Fold_Change = sign(logFC)*2^abs(logFC), .keep = "all")
  
  fullMatrix <- left_join((geneAnnot %>% dplyr::select(-pl_length, -pl_gc, -is_principal,
                                                       -GENEBIOTYPE, -tx_id, -gene_id_version)),
                          (dge$counts %>% as.data.frame() %>% rownames_to_column("gene_id"))
  )
  
  cols <- c('gene_id', 'SYMBOL', 'DESCRIPTION', 'Fold_Change', 'logFC', 'PValue',
            'FDR', sprintf("%s_Avg_FPKM", condition1), sprintf("%s_Avg_FPKM", condition2))
  
  presentMatrix <- diffExpGenes %>% dplyr::select(cols)
  
  ssMatrix <- statSigFilter(diffExpGenes, condition1, condition2)
  
  bsMatrix <- bioSigFilter(diffExpGenes, condition1, condition2)
  
  resList <- list(
    `Sample Title` = degTitle,
    `Sample Comparison` = sampleComparison,
    `Contrast Condition` = contrastCondition,
    `Condition 1` = condition1,
    `Condition 2` = condition2,
    `Full Count Matrix` = fullMatrix,
    `All Present Genes` = presentMatrix,
    `Statistically Significant` = ssMatrix,
    `Biologically Significant` = bsMatrix
  )
    
  return(resList)
}


statSigFilter <- function(deg, condition1, condition2){
  cols = c(
    'gene_id', 'SYMBOL', 'DESCRIPTION', 'Fold_Change', 'logFC', 'PValue',
    'FDR', sprintf("%s_Avg_FPKM", condition1), sprintf("%s_Avg_FPKM", condition2)
  )
  ssDeg <- deg %>% 
    filter(abs(logFC) > 0 & FDR < 0.05) %>%
    dplyr::select(cols)
  
  row.names(ssDeg) <- 1:nrow(ssDeg)
  return(ssDeg)
}

bioSigFilter <- function(deg, condition1, condition2){
  cols = c(
    'gene_id', 'SYMBOL', 'DESCRIPTION', 'Fold_Change', 'logFC', 'PValue',
    'FDR', "cKO_Avg_FPKM", "WT_Avg_FPKM"
  )
  bsDeg <- deg %>%
    filter(abs(logFC) > 1 & FDR < 0.05) %>%
    filter(abs(WT_Avg_FPKM - cKO_Avg_FPKM) > 2) %>%
    filter(abs(WT_Avg_FPKM) > 2 | abs(cKO_Avg_FPKM) > 2) %>%
    dplyr::select(cols)
  
  row.names(bsDeg) <- 1:nrow(bsDeg)
  return(bsDeg)
}

genDegOverlap <- function(degobj1="deg1", #First deg object to compare
                          degobj2="deg2" #Second deg object to compare
                          ){
  
  deg1 <- degobj1$`All Present Genes`; deg2 <- degobj2$`All Present Genes`
  
  interDeg<-full_join(x=deg1, y=deg2, 
                      by=c("gene_id","SYMBOL","DESCRIPTION"), 
                      suffix=paste("_",c(degobj1$`Contrast Condition`,degobj2$`Contrast Condition`),sep=""))
  
  interSsDeg <- full_join(x=degobj1$`Statistically Significant`, 
                          y = degobj2$`Statistically Significant`,
                          by=c("gene_id","SYMBOL","DESCRIPTION"),
                          suffix=paste("_",c(degobj1$`Contrast Condition`,degobj2$`Contrast Condition`),sep=""))
  
  interBsDeg <- full_join(x=degobj1$`Biologically Significant`,
                          y = degobj2$`Biologically Significant`,
                       by=c("gene_id","SYMBOL","DESCRIPTION"),
                       suffix=paste("_",c(degobj1$`Contrast Condition`,degobj2$`Contrast Condition`),sep=""))
  
  deg1.lfc <- as.symbol(paste("logFC", degobj1$`Contrast Condition`, sep = "_"))
  deg2.lfc <- as.symbol(paste("logFC", degobj2$`Contrast Condition`, sep = "_"))
  
  subSsInterList <- list(
    interSsDeg %>% filter(!!deg1.lfc > 0 & !!deg2.lfc > 0 ),
    interSsDeg %>% filter(!!deg1.lfc > 0 & !!deg2.lfc < 0 ),
    interSsDeg %>% filter(!!deg1.lfc < 0 & !!deg2.lfc > 0 ),
    interSsDeg %>% filter(!!deg1.lfc < 0 & !!deg2.lfc < 0 )
  )
  
  names(subSsInterList) <- c(sprintf("SS Up %s Up %s", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                             sprintf("SS Up %s Dn %s", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                             sprintf("SS Dn %s Up %s", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                             sprintf("SS Dn %s Dn %s", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`)
                             )
  subSsWithNA <- list(
    interSsDeg %>% filter(!!deg1.lfc > 0 & is.na(!!deg2.lfc)),
    interSsDeg %>% filter(!!deg1.lfc < 0 & is.na(!!deg2.lfc)),
    interSsDeg %>% filter(is.na(!!deg1.lfc) & !!deg2.lfc > 0),
    interSsDeg %>% filter(is.na(!!deg1.lfc) & !!deg2.lfc < 0)
  )
  
  names(subSsWithNA) <- c(sprintf("SS Up %s NoExp %s", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                          sprintf("SS Dn %s NoExp %s", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                          sprintf("SS NoExp %s Up %s", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                          sprintf("SS NoExp %s Dn %s", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`)
                          )
  
  
  subBsInterList <- list(
    interBsDeg %>% filter(!!deg1.lfc > 0 & !!deg2.lfc > 0 ),
    interBsDeg %>% filter(!!deg1.lfc > 0 & !!deg2.lfc < 0 ),
    interBsDeg %>% filter(!!deg1.lfc < 0 & !!deg2.lfc > 0 ),
    interBsDeg %>% filter(!!deg1.lfc < 0 & !!deg2.lfc < 0 )
  )
  
  names(subBsInterList) <- c(sprintf("BS Up %s Up %s", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                             sprintf("BS Up %s Dn %s", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                             sprintf("BS Dn %s Up %s", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                             sprintf("BS Dn %s Dn %s", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`)
                             )
  
  subBsWithNA <- list(
    interBsDeg %>% filter(!!deg1.lfc > 0 & is.na(!!deg2.lfc)),
    interBsDeg %>% filter(!!deg1.lfc < 0 & is.na(!!deg2.lfc)),
    interBsDeg %>% filter(is.na(!!deg1.lfc) & !!deg2.lfc > 0),
    interBsDeg %>% filter(is.na(!!deg1.lfc) & !!deg2.lfc < 0)
  )
  
  names(subBsWithNA) <- c(sprintf("BS Up %s NoExp %s", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                          sprintf("BS Dn %s NoExp %s", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                          sprintf("BS NoExp %s Up %s", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`),
                          sprintf("BS NoExp %s Dn %s", degobj1$`Contrast Condition`, degobj2$`Contrast Condition`)
  )
  
  resList <- list(interSsDeg, subSsInterList, subSsWithNA,
                  interBsDeg, subBsInterList, subBsWithNA,
                  interDeg
                 )
  
  names(resList) <- c("Stat Sig Intersection","Subset SS", "Subset SS w/ NA", 
                      "Bio Sig Intersection", "Subset BS", "Subset BS w/ NA",
                      "Full Intersection"
                       )
  
  
  return(resList)
}