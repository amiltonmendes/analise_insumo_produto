# Limpandoo ambiente
rm(list = ls())

# =========== Área de bibliotecas =================
# Pacotes necessários

if (!require("httr")) install.packages("httr")
if (!require("readxl")) install.packages("readxl")
if (!require("writexl")) install.packages("writexl")
if (!require("fs")) install.packages("fs")

library(httr)
library(readxl)
library(writexl)
library(fs)
# =========== Área de Funções =====================



# =========== Área do código principal ============
#
# Download das Tabela sde REcursos e Usos
#
# URL do ZIP com as Tabelas de Recursos e Usos (ajuste conforme o ano desejado)
url <- "https://ftp.ibge.gov.br/Contas_Nacionais/Sistema_de_Contas_Nacionais/2021/tabelas_xls/tabelas_de_recursos_e_usos/nivel_68_2010_2021_xls.zip"
#
# Nome do arquivo para salvar localmente
ArquivoDestino <- "Recursos_Usos_2021.zip"

# Diretório de destino
PastaDestino <- "./TRU_2021"
dir_create(PastaDestino)

# Download do arquivo ZIP
GET(url, write_disk(file.path(PastaDestino,ArquivoDestino), overwrite = TRUE))

# Extração do ZIP
unzip(file.path(PastaDestino,ArquivoDestino),
      files = NULL,
      overwrite = TRUE,
      exdir = PastaDestino)
#
# Lendo as Tabelas
nSetores = 68
nProdutos = 128
# Lendo Oferta
dSheet <- read_excel(file.path(PastaDestino,"68_tab1_2019.xls"),
                     sheet = "oferta")
mMatriz = dSheet[5:132, 3:9]
mOferta= apply(as.matrix.noquote(mMatriz),2,as.numeric)

# lendo Producao
dSheet <- read_excel(file.path(PastaDestino,"68_tab1_2019.xls"),
                     sheet = "producao")
mMatriz = dSheet[5:132, 3:70]
mProducao= apply(as.matrix.noquote(mMatriz),2,as.numeric)

# Lendo Importação
dSheet <- read_excel(file.path(PastaDestino,"68_tab1_2019.xls"),
                     sheet = "importacao")
mMatriz = dSheet[5:132, 3]
vImportacao= apply(as.matrix(mMatriz),2,as.numeric)

#lendo Consumo Intermediário
dSheet <- read_excel(file.path(PastaDestino,"68_tab2_2019.xls"),
                     sheet = "CI")
mMatriz = dSheet[5:132, 3:70]
mCI= apply(as.matrix.noquote(mMatriz),2,as.numeric)

#lendo demanda
dSheet <- read_excel(file.path(PastaDestino,"68_tab2_2019.xls"),
                     sheet = "demanda")
mMatriz = dSheet[5:132, 3:8]
mDemanda= apply(as.matrix.noquote(mMatriz),2,as.numeric)

#lendo VA
dSheet <- read_excel(file.path(PastaDestino,"68_tab2_2019.xls"),
                     sheet = "VA")
mMatriz = dSheet[5:18, 2:69]
mVA= apply(as.matrix.noquote(mMatriz),2,as.numeric)





# Calculando Matriz e distribuição sem Variação de estoque
nColunaEstoque = 6
mDemandaFinalSemEstoque <- mDemanda
mDemandaFinalSemEstoque[, nColunaEstoque] <- 0.0
mConsumoTotalSemEstoque <- cbind(mCI, mDemandaFinalSemEstoque)

vTotalProduto = rowSums(mConsumoTotalSemEstoque)
mDistribuicao = mConsumoTotalSemEstoque / vTotalProduto
mDistribuicao[is.na(mDistribuicao)] <- 0

# Distribuir  IPI, ICMS e OILL
nColunaIPI = 5 
mValorIPI = mOferta[,nColunaIPI] * mDistribuicao
nColunaICMS = 6
mValorICMS = mOferta[,nColunaICMS] * mDistribuicao
nColunaOILL = 7
mValorOILL = mOferta[,nColunaOILL] * mDistribuicao



# Distribui a margem  do comércio 
nColunaMargemComercio = 2
nColunaMargemTransporte = 3

vVetorEntrada <-mOferta[, nColunaMargemComercio]
vMargem <- c(93, 94)

vPropMargem <- vVetorEntrada[vMargem[1]:vMargem[2]] / sum(vVetorEntrada[vMargem[1]:vMargem[2]])
mMargemComercio <- vVetorEntrada * mDistribuicao

mMargemDistribuida <- colSums(mMargemComercio[1:(vMargem[1]-1), ]) + colSums(mMargemComercio[(vMargem[2]+1):nrow(mMargemComercio), ])

mMargemComercio[vMargem[1]:vMargem[2], ] <- t(t(vPropMargem))  %*% mMargemDistribuida * (-1)
options(scipen = 0)
colSums(mMargemComercio)

# Distribui a margem  do transporte 

vVetorEntrada <-mOferta[, nColunaMargemTransporte]
vMargem <- c(95, 98)

vPropMargem <- vVetorEntrada[vMargem[1]:vMargem[2]] / sum(vVetorEntrada[vMargem[1]:vMargem[2]])
mMargemTransporte <- vVetorEntrada * mDistribuicao
mMargemDistribuida <- colSums(mMargemTransporte[1:(vMargem[1]-1), ]) + colSums(mMargemTransporte[(vMargem[2]+1):nrow(mMargemTransporte), ])
mMargemTransporte[vMargem[1]:vMargem[2], ] <- t(t(vPropMargem))  %*% mMargemDistribuida * (-1)


# Calcula matriz de distribuição sem exportação
nColunaExportacao <- 1
mDemandaFinalSemExportacao <- mDemandaFinalSemEstoque
mDemandaFinalSemExportacao[, nColunaExportacao] <- 0.0
mConsumoTotalSemExportacao <- cbind(mCI, mDemandaFinalSemExportacao)

vTotalProduto <- rowSums(mConsumoTotalSemExportacao)
mDistribuicaoImportacao <- mConsumoTotalSemExportacao / vTotalProduto
mDistribuicaoImportacao[is.na(mDistribuicaoImportacao)] <- 0

# Distribuir  II e importação
nColunaII <- 4
mValorII <- mOferta[, nColunaII] * mDistribuicaoImportacao
mImportacao <- vImportacao[,] * mDistribuicaoImportacao


# Calcula o consumo total a preço basicos
mConsumoTotalPrecoMercado <- cbind(mCI, mDemanda)
mConsumoTotalPrecoBase <- mConsumoTotalPrecoMercado - mMargemComercio - mMargemTransporte - mValorIPI - mValorICMS - mValorOILL - mImportacao - mValorII



# Estimação da MIP setor x setor utilizando a tecnologia baseada na indústria

nLinhasVA <- 14
nLinhaVBP <- 13
nLinhaOcupacoes <- 14
vVBP <- mVA[nLinhaVBP,]
vOcupacoes <- mVA[nLinhaOcupacoes,]

mU =mConsumoTotalPrecoBase[,1:nSetores]
mE =mConsumoTotalPrecoBase[,(nSetores+1):ncol(mConsumoTotalPrecoBase)]
options(scipen = 2)

vX = mVA[nLinhaVBP,]
mXChapeu = diag(1/vX)

mB = mU %*% mXChapeu

mV = t(mProducao)
vQ = colSums(mV)
mQChapeu = diag(1/vQ)

mD = mV %*% mQChapeu

mA = mD %*% mB
mY = mD %*% mE
mZ = mD %*% mU

mI = diag(nSetores)

mLeontief = (mI - mA)

vDemandaTotal = rowSums(mZ) + rowSums(mY)
nDemandaTotal = sum(vDemandaTotal)

mTConsumoItermediario = rbind(colSums(mValorII[,1:nSetores]),
                              colSums(mValorIPI[,1:nSetores]),
                              colSums(mValorICMS[,1:nSetores]),
                              colSums(mValorOILL[,1:nSetores]) )

nLinhaVA=1
vOfertaTotal = colSums(mZ) + colSums(mImportacao[,1:nSetores]) +
  colSums(mTConsumoItermediario) + mVA[nLinhaVA,]    
nOfertaTotal = sum(vOfertaTotal)
vDiferencas = vDemandaTotal - vOfertaTotal
vDiferencaTotal = sum(vDiferencas)
vDiferencas

# Montagem da Matriz - Preparação dos totais


vMConsumoIntermdiario = colSums(mImportacao[,1:nSetores])
vMDemandaFinal        = colSums(mImportacao[,nSetores+1:6])

vValorIIConsumoIntermdiario = colSums(mValorII[,1:nSetores])
vValorIIDemandaFinal        = colSums(mValorII[,nSetores+1:6])

vValorIPIConsumoIntermdiario = colSums(mValorIPI[,1:nSetores])
vValorIPIDemandaFinal        = colSums(mValorIPI[,nSetores+1:6])

vValorICMSConsumoIntermdiario = colSums(mValorICMS[,1:nSetores])
vValorICMSDemandaFinal        = colSums(mValorICMS[,nSetores+1:6])

vValorOIConsumoIntermdiario = colSums(mValorOILL[,1:nSetores])
vValorOIDemandaFinal        = colSums(mValorOILL[,nSetores+1:6])

mTConsumoIntermediario = rbind(vValorIIConsumoIntermdiario, vValorIPIConsumoIntermdiario, 
                               vValorICMSConsumoIntermdiario, vValorOIConsumoIntermdiario)

mTDemandaFinal = rbind(vValorIIDemandaFinal, vValorIPIDemandaFinal, 
                       vValorICMSDemandaFinal, vValorOIDemandaFinal)

# Montagem da Oferta
vOfertaNacional <- colSums(mZ)
vOfertaTributos <- colSums(mTConsumoIntermediario)
vOfertaConsumoIntermediario <- vOfertaNacional + vMConsumoIntermdiario + vOfertaTributos
mOfertaConsumoIntermediario <- rbind(mZ, vOfertaNacional, vMConsumoIntermdiario, mTConsumoIntermediario,
                                     vOfertaConsumoIntermediario, mVA)

# Calculando Demanda Intermediaria Total
mTotalDemandaConsumoIntermediario <- rowSums(mOfertaConsumoIntermediario)

# Montagem da Demanda Final

mTDemandaFinal <- rbind(vValorIIDemandaFinal, vValorIPIDemandaFinal, vValorICMSDemandaFinal, vValorOIDemandaFinal)
vTotalLinhaDemandaNacional <- colSums(mY)
mDemanda <- rbind(mY, vTotalLinhaDemandaNacional, vMDemandaFinal, mTDemandaFinal)
vTotalLinhaDemandaFinal <- colSums(mDemanda)
mDemandaFinal <- rbind(mDemanda, vTotalLinhaDemandaFinal, matrix(0, nrow = 14, ncol = 6))

vTotalDemandaFinal <- rowSums(mDemandaFinal)
vDemandaTotal <- mTotalDemandaConsumoIntermediario + vTotalDemandaFinal
mMIP <- cbind(mOfertaConsumoIntermediario, mTotalDemandaConsumoIntermediario, mDemandaFinal,
              vTotalDemandaFinal, vDemandaTotal)


# lendo Nome de Setores
lNomeSetores <- c()
for (i in 1:nSetores) {
  lNomeSetores[[i]] <- dSheet[3, 1 + i]
  lNomeSetores[[i]] <- substr(lNomeSetores[[i]],6,50)
}

lNomeVA = c()
for (i in 1:14) {
  lNomeVA[[i]] <- dSheet[3+i, 1]
}

sColsLabel <- c()
sRowsLabel <- c()

sRowsLabel <- c(lNomeSetores, "Oferta Nacional", "Importação", "II", "IPI", "ICMS", "OI", "Oferta Intermediária Total", lNomeVA)

sColsLabel <- c(lNomeSetores, "Consumo Intermediário", "Exportação", "Consumo do Governo", "Consumo das ISFLSF", "Consumo das Famílias", "FBCF", "Variação do estoque", "Demanda Final", "Demanda Total")

# Definindo o caminho de saída
nAnoMIP=2019
PastaDestino <- "./TRU_2021"

sCaminhoOutput <- "Outputs/"
dir_create(sCaminhoOutput)
nome_arquivo <- paste0("MIP_", as.character(nAnoMIP), "_", as.character(nSetores), "x", as.character(nSetores), "MetodoGuilhoto.xlsx")



# Converter a matriz 'mMIP' em um data frame
df <- as.data.frame(mMIP)
rownames(df) <- sRowsLabel
colnames(df) <- sColsLabel

# Escrever no arquivo Excel
()
write_xlsx(df, paste0(sCaminhoOutput, nome_arquivo))
print("Acabou")

