# Limpandoo ambiente
rm(list = ls())

# =========== Área de bibliotecas =================
# Pacotes necessários
if (!require("httr")) install.packages("httr")
if (!require("openxlsx")) install.packages("openxlsx")
if (!require("fs")) install.packages("fs")
if (!require("htmltools")) install.packages("htmltools")
if (!require("htmlTable")) install.packages("htmlTable")


library(httr)
library(openxlsx)
library(fs)
library(htmltools)
library(htmlTable)
# =========== Área de Funções =====================



# =========== Área do código principal ============
#
# Leitura da Matriz
#
nSetores = 68
nProdutos = 128
nColunaFinal= 77
nLInhaFinal = 90
nNumColunasDemanda = 6

PastaMipEstimada <- "./Outputs"
PastaMipEstimada <- fs::path_expand(PastaMipEstimada)
dSheet <- read.xlsx(file.path(PastaMipEstimada,"MIP_2019_68x68MetodoGuilhoto.xlsx"),
                                 sheet = "MIP",
                                 rowNames = TRUE)
mMatriz  <- as.matrix(dSheet)
dimnames(mMatriz) <- NULL

mZ <- mMatriz[1:nSetores, 1:nSetores]
mDemanda  <- mMatriz[1:nSetores, 70:75]
vMCI      <- mMatriz[70, 1:nSetores]
vMdemanda <- mMatriz[70, 70:75]
mTCI      <- mMatriz[71:74, 1:nSetores]
vTDemanda <- mMatriz[71:74, 70:75]
mVA       <- mMatriz[76:89, 1:nSetores]
nLinhasVA <- 14
nLinhaVBP <- 13
nLinhaOcupacoes <- 14
vVBP      <- mVA[nLinhaVBP,]
vX        <- mVA[nLinhaVBP,]
vOcupacoes<- mVA[nLinhaOcupacoes,]

mA <- mZ / vX

mI <- diag(nSetores)

mLeontief <- solve(mI - mA)


#################################################################################
# Choque no Modelo Aberto
#################################################################################
# Leitura do Choque

vChoque <- matrix(0, nrow = nSetores, ncol = 1)
mB <- mLeontief 
#vChoque[, 1] <- 0.1
#nColunaExportacao <- 1  # Substitua pelo valor correto de nColunaExportacao
#vExportacao <- matrix(mY[, nColunaExportacao], nrow = nSetores, ncol = 1)

#vDeltaY <- vChoque * vExportacao
vChoque[1,1]= 10000
vDeltaY=vChoque

nVBP = sum(vX)
vDeltaX <- mB %*% vDeltaY
nVariacaoVBP <- sum(vDeltaX)
nVariacaoPercentualVBP <- sum(vDeltaX) / nVBP * 100

vRelacaoVA <- (mVA[1,] / vX)  
vDeltaVA <- vDeltaX * vRelacaoVA
nDeltaVA <- sum(vDeltaX * vRelacaoVA)
nDeltaVAPercentual <- nDeltaVA / sum(mVA[1,]) * 100

vRelacaoEmprego <- (mVA[nLinhaOcupacoes,]  / vX)


vDeltaEmprego <- vDeltaX * vRelacaoEmprego 
nDeltaEmprego = sum(vDeltaEmprego)
nDeltaEmpregoPercentual = nDeltaEmprego / sum(mVA[nLinhaOcupacoes,])  * 100

resumo <- data.frame(
  Indicador = c("Valor Total da Produção (R$)", 
                "Variação da Produção (R$)", 
                "Variação da Produção (%)",
                "Variação do VA - R$ ", 
                "Variação do VA - %  ",
                "Variação do emprego -  ",  
                "Variação do emprego - %  "
                ),
  Valor = c(format(nVBP, big.mark = ".", decimal.mark = ",", digits = 2), 
            format(nVariacaoVBP, big.mark = ".", decimal.mark = ",", digits = 2), 
            paste0(format(nVariacaoPercentualVBP, digits = 2), " %"), 
            format(nDeltaVA, big.mark = ".", decimal.mark = ",", digits = 2),
            paste0(format(nDeltaVAPercentual, digits = 2), " %"),
            format(trunc(nDeltaEmprego), big.mark = ".", decimal.mark = ",", digits = 2),
            paste0(format(nDeltaEmpregoPercentual, digits = 2), " %")
            )
)

html_out <- htmlTable(resumo, align = "l", caption = "Choque no Modelo Aberto")
html_print(HTML(html_out))

#################################################################################
# Calculo do Modelo Fechado
#################################################################################
nLinhaRendaTrabalho <- 2
nColunaFamilias <- 4
vDemandaFamilias <- matrix(mDemanda[, nColunaFamilias], nrow = nSetores, ncol = 1)
vRendaTrabalho <- matrix(mVA[nLinhaRendaTrabalho, ], nrow = 1, ncol = nSetores) + matrix(mVA[9, ], nrow = 1, ncol = nSetores) * 0.41
vAuxZero <- matrix(0, nrow = 1, ncol = 1)

mZBarr <- rbind(mZ, vRendaTrabalho)
mAux <- rbind(vDemandaFamilias, vAuxZero)
mZBarr <- cbind(mZBarr, mAux)

vAuxZero[1, 1] <- sum(vX)
vVBPBarr <- cbind(matrix(vX, nrow = 1, ncol = nSetores), vAuxZero)

mABarr <- matrix(0, nrow = nSetores + 1, ncol = nSetores + 1)
#### mABarr[,] <- mZBarr[,] / vVBPBarr[1, ]
for (i in 0:nSetores+1) {
  mABarr[i,] <- mZBarr[i,] / vVBPBarr[1,]
}
mIBarr <- diag(nSetores + 1)
mLeontiefBarr <- solve(mIBarr - mABarr)

#################################################################################
# Choque no Modelo Fechado
#################################################################################

vChoque <- matrix(0, nrow = nSetores, ncol = 1)
mBBarr  = mLeontiefBarr
vChoque[, 1] <- 0.1
nColunaExportacao <- 1  # Substitua pelo valor correto de nColunaExportacao
vExportacao <- matrix(mDemanda[, nColunaExportacao], nrow = nSetores, ncol = 1)

vDeltaY <- vChoque * vExportacao
vAuxZero <- matrix(0, nrow = 1, ncol = 1)

vDeltaYBarr <- rbind(vDeltaY, vAuxZero)

nVBP = sum(vX)
vDeltaXBarr <- mLeontiefBarr %*% vDeltaYBarr
nVariacaoVBPBarr <- sum(vDeltaXBarr)
nVariacaoPercentualVBPBarr <- sum(vDeltaXBarr[1:nSetores]) / nVBP * 100

vRelacaoVA <- (mVA[1,] / vX)  
vDeltaVABarr <- vDeltaXBarr[1:nSetores] * vRelacaoVA
nDeltaVABarr <- sum(vDeltaVABarr)
nDeltaVAPercentualBarr <- nDeltaVABarr / sum(mVA[1,]) * 100
vRelacaoEmprego <- (mVA[nLinhaOcupacoes,]  / vX)
vDeltaEmprego <- vDeltaX * vRelacaoEmprego 

resumo <- data.frame(
  Indicador = c("Valor Total da Produção (R$)", 
                "Variação da Produção (R$)", 
                "Variação da Produção (%)",
                "Variação do VA - R$ ", 
                "Variação do VA - %  ",
                "Variação do emprego -  ",  
                "Variação do emprego - %  "
  ),
  Valor = c(format(nVBP, big.mark = ".", decimal.mark = ",", digits = 2), 
            format(nVariacaoVBPBarr, big.mark = ".", decimal.mark = ",", digits = 2), 
            paste0(format(nVariacaoPercentualVBPBarr, digits = 2), " %"), 
            format(nDeltaVABarr, big.mark = ".", decimal.mark = ",", digits = 2),
            paste0(format(nDeltaVAPercentualBarr, digits = 2), " %"),
            format(trunc(nDeltaEmprego), big.mark = ".", decimal.mark = ",", digits = 2),
            paste0(format(nDeltaEmpregoPercentual, digits = 2), " %")
  )
)

html_out <- htmlTable(resumo, align = "l", caption = "Choque no Modelo Fechado")
html_print(HTML(html_out))






#################################################################################
# Calculando Multiplicadores de Impacto
# Multiplicador de produção
#################################################################################
vMultTotalProducao <- colSums(mLeontiefBarr)
vMultSimplesProducao <- colSums(mLeontief)
vMultTotalProducaoTrunc <- colSums(mLeontiefBarr[1:nSetores, 1:nSetores])
vEfeitoInduzido <- vMultTotalProducao[1:nSetores] - vMultSimplesProducao
vEfeitoDireto <- colSums(mA[1:nSetores, 1:nSetores])
vEfeitoIndireto <- vMultSimplesProducao - vEfeitoDireto
if (!require(stargazer)) {
  install.packages("stargazer")
  library(stargazer)
}
library(stargazer)

mMultiplicadoresProducao <- cbind(vMultTotalProducao[1:nSetores], vMultSimplesProducao, vEfeitoInduzido, vEfeitoDireto, vEfeitoIndireto)

colnames(mMultiplicadoresProducao) <- c("Total", "Simples", "Induzido", "Direto", "Indireto")

rownames(mMultiplicadoresProducao) = c(lNomeSetores)

htmlFile <- file.path("tempfile.html")
stargazer::stargazer(mMultiplicadoresProducao, type = "html", header = FALSE, out = htmlFile)

rstudioapi::viewer(htmlFile)


#################################################################################
# Calculando Multiplicadores de Impacto
# Multiplicador de Emprego
#################################################################################
vOcupacoes = mVA[nLinhaOcupacoes,]
vRequisitosEmprego <- vOcupacoes /  vX
mRequisitosEmpregoDiagonal <- diag(vRequisitosEmprego)
mGeradorEmprego <- mRequisitosEmpregoDiagonal %*% mLeontief
vMultSimplesEmprego <- colSums(mGeradorEmprego)
vMultiplicadorEmpregoI <- vMultSimplesEmprego / vRequisitosEmprego

mGeradorEmpregoModeloFechado <- mRequisitosEmpregoDiagonal %*% mLeontiefBarr[1:nSetores, 1:nSetores]
vMultTotalEmprego <- colSums(mGeradorEmpregoModeloFechado)
vMultiplicadorEmpregoII <- vMultTotalEmprego / vRequisitosEmprego

vEfeitoInduzidoEmprego <- vMultTotalEmprego - vMultSimplesEmprego
vEfeitoDiretoEmprego <- vRequisitosEmprego
vEfeitoIndiretoEmprego <- vMultSimplesEmprego - vEfeitoDiretoEmprego



mMultiplicadoresEmprego <- cbind(vMultTotalEmprego, vMultSimplesEmprego, vEfeitoInduzidoEmprego, vEfeitoDiretoEmprego, vEfeitoIndiretoEmprego)

colnames(mMultiplicadoresEmprego) <- c("Total", "Simples", "Induzido", "Direto", "Indireto")

for (i in 1:nSetores) {
  rownames(mMultiplicadoresEmprego)[i] <- substr(lNomeSetores[i], 1, 60)
}

htmlFile <- file.path("tempfile.html")
stargazer::stargazer(mMultiplicadoresEmprego, type = "html", header = FALSE, out = htmlFile)

rstudioapi::viewer(htmlFile)



#################################################################################
# Calculando Multiplicadores de Impacto
# Multiplicador de Renda
#################################################################################
nLinhaRenda = 1
vRequisitosRenda <- mVA[nLinhaRenda, ] / vX
mRequisitosRendaDiagonal <- diag(vRequisitosRenda)
mGeradorRenda <- mRequisitosRendaDiagonal %*% mLeontief
vMultSimplesRenda <- colSums(mGeradorRenda)
vMultiplicadorRendaI <- vMultSimplesRenda / vRequisitosRenda

mGeradorRendaModeloFechado <- mRequisitosRendaDiagonal %*% mLeontiefBarr[1:nSetores, 1:nSetores]
vMultTotalRenda <- colSums(mGeradorRendaModeloFechado)
vMultiplicadorRendaII <- vMultTotalRenda / vRequisitosRenda

vEfeitoInduzidoRenda <- vMultTotalRenda - vMultSimplesRenda
vEfeitoDiretoRenda <- vRequisitosRenda
vEfeitoIndiretoRenda <- vMultSimplesRenda - vEfeitoDiretoRenda

mMultiplicadoresRenda <- cbind(vMultTotalRenda, vMultSimplesRenda, vEfeitoInduzidoRenda, vEfeitoDiretoRenda, vEfeitoIndiretoRenda)

colnames(mMultiplicadoresRenda) <- c("Total", "Simples", "Induzido", "Direto", "Indireto")


for (i in 1:nSetores) {
  rownames(mMultiplicadoresRenda)[i] <- substr(lNomeSetores[i], 1, 60)
}
library(tidyverse)
htmlFile <- file.path("tempfile.html")
stargazer::stargazer(mMultiplicadoresRenda, type = "html", header = FALSE, out = htmlFile)

rstudioapi::viewer(htmlFile)


