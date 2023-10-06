#==================================================================
#==================================================================
#================== CALCULO DA DISTRIBUICAO CMPC (CMPC)  =================
#==================================================================
#==================================================================

# O presente algoritmo foi desenvolvido pela ANTT com o unico proposito
# de permitir a atualizacao transparente da distribuição de probabilidade 
# do custo medio ponderado de capital (CMPC) para o setor FERROVIÁRIO

#   ENTRADA (INPUT):   Planilha "Variaveis_CMPC.xlsx"

#   SAIDA (OUTPUT):    Planilha "Resultado_CMPC.xlsx"

#  O resultado consiste no CMPC para cada percentil, de 1% até 100% da distribuicao obtida


#==================================================================
#============= INSTALACAO E CARREGAMENTO DE PACOTES  ==============
#==================================================================

#---------------- Instalacao de pacotes necessarios ----------------


list.of.packages <- c("EnvStats", "freedom","BAT","readxl","writexl","rstudioapi","here","fitdistrplus","mc2d")
new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
if(length(new.packages)) install.packages(new.packages)


#---------------------------VARIÁVEIS PARA A SIMULAÇÃO---------------

#lambda_desativado = sample(c(TRUE, FALSE), 1)
#corrige_janela_PRM = sample(c(TRUE, FALSE), 1)
#riscos_MF = sample(c(TRUE, FALSE), 1)
#Rd_Bacen = sample(c(TRUE, FALSE), 1)

lambda_desativado = TRUE
corrige_janela_PRM = FALSE
riscos_MF = FALSE
Rd_Bacen = FALSE

opcoes = c("10PRM30", "10", "20", "30")
Anos = sample(opcoes, 1)

#---------------------------VARIÁVEIS PARA A SIMULAÇÃO---------------


#---------------- Carregando pacotes necessarios ------------------
  
  library(EnvStats)
  library(freedom)
  library(BAT)
  library(readxl)
  library(writexl)
  library(rstudioapi)
  library(here)
  library(fitdistrplus)
  library(mc2d)


# Mudança do diretorio de trabalho para a pasta onde esta o script  
 caminho <- dirname(rstudioapi::getSourceEditorContext()$path)
 setwd(here(caminho))
 
#==================================================================
#======================= IMPORTACAO DE DADOS ======================
#==================================================================

#--------------- Selecao do arquivo "Variaveis CMPC" ------------

 Dados <- paste(caminho, "/Variaveis_CMPC_adaptado_PERIODOS", ".xlsx",sep="")
 

#-------- Leitura para o R dos dados inseridos no arquivo "Variaveis_CMPC" -------  

 BCB_cred<- read_excel(Dados, sheet = "Rd_alternativo") #Dados de Crédito do Banco Central
 
 PRM.1<- read_excel(Dados, sheet = "PRM.1") #Premio de risco de mercado (PRM)
 PRM.2<- read_excel(Dados, sheet = "PRM.2") #Premio de risco de mercado (PRM)
 IPCA<-read_excel(Dados, sheet = "IPCA") # Inflacao (IPCA)
 CDI<- read_excel(Dados, sheet = "CDI") # Benchmaking financiamento bancario (CDI)
 TLP_pre<- read_excel(Dados, sheet = "TLP pre") # Benchmaking financiamento BNDES (TLP)
 FONTES_FIN<- read_excel(Dados, sheet = "Fontes_financiamento") #Proporcao das fontes de financiamento
 EST_CAP <- read_excel(Dados, sheet = "Estrutura_capital") #Estrutura de Capital (EST_CAP)
 BETA <- read_excel(Dados, sheet = "Beta") # Beta (BETA)
 DEB_IPCA <- read_excel(Dados, sheet = "Deb_IPCA") # informacoes de custo de captacao de debentures IPCA+
 DEB_DI <- read_excel(Dados, sheet = "Deb_DI") # informacoes de custo de captacao de debentures CDI+
 CST_BNDES <- read_excel(Dados, sheet = "Custo_BNDES") # informacoes de custo de captação junto ao BNDES
 EF_TRIBUT <- read_excel(Dados, sheet = "Efeito_Trib") # Aliquota de imposto de renda e contribuicao social
 ITER <- read_excel(Dados, sheet = "Iteracoes") # Quantidade de iteracoes das simulacoes
 idka5 <-read_excel(Dados, sheet = "IDKA5") #
 
 PRM.3 <- read_excel(Dados, sheet = "PRM.3")
 
 RF30 = PRM.3$RF 
 RF20 = tail(PRM.3$RF , 240)
 
 RP30 = PRM.3$RP #d[!is.na(d)]
 RP30 = RP30[!is.na(RP30)]
 RP20 = tail(PRM.3$RP , 240)
 
 CPI30 = PRM.3$CPI
 CPI20 = tail(PRM.3$CPI , 240)
 
 PRM20 <- tail(PRM.2$PRM, 240)
 #Atribuicao dos conjuntos de valores para as variaveis aleatorias utilizadas no algoritmo

 DIRE<- BCB_cred$dir #(10 anos)
 TPB<- BCB_cred$tpb #(10 anos)
 
 
 if (Anos == "10PRM30") { 
   PRM <- PRM.2$PRM
   RF <- PRM.1$RF
   RP<- PRM.1$RP
   CPI <- PRM.1$CPI
 } else if (Anos == "10") {
   PRM <- tail(PRM.2$PRM, 120)
   RF <- PRM.1$RF
   RP<- PRM.1$RP
   CPI <- PRM.1$CPI
 } else if  (Anos == "20") {
   PRM <- PRM20
   RF <- RF20
   RP<-  RP20
   CPI <- CPI20
 } else {
   PRM <- PRM.2$PRM
   RF <- RF30
   RP<- RP30
   CPI <- CPI30
 }
 
 

 IPCA<- IPCA$ipca #(10 anos)
 CDI<- CDI$cdi
 TLP_pre<- TLP_pre$tlp_pre
 idka5 <-idka5$idka5
 
 

#==================================================================
#===================== SIMULACOES MONTE CARLO =====================
#==================================================================
 
#----------------------------- FITTING ----------------------------
  
#-------- TESTE DE ADERENCIA DAS DISTRIBUICOES PARAMETRICAS -------  
  
  #Teste de AIC (selecionar melhor distribuicao parametrica)
  

  variaveis <- c('RP', 'IPCA','CDI','TLP_pre','CPI','RF','PRM','idka5', 'DIRE', 'TPB')
  variavel_list <- list('RP'=RP,'IPCA'=IPCA,'CDI'=CDI,'TLP_pre'=TLP_pre,'CPI'=CPI,'RF'=RF,'PRM'=PRM,'idka5'=idka5, 'DIRE'=DIRE, 'TPB'=TPB)

  for(i in variaveis){

      #FITTING Distribuicao normal
        fit_normal <- fitdist(unlist(variavel_list[i]),"norm", method="mge")
        parametro <- unlist(fit_normal['estimate'])
        media_var <- as.numeric(parametro[1])
        desv.p_var <- as.numeric(parametro[2])
        dist.norm <-   rnorm(ITER$Qtd_iteracoes, media_var, desv.p_var)
        aic_norm <- aic(unlist(variavel_list[i]), dist.norm)
      
      #FITTING Distribuicao triangular
      
        fit_tri <- fitdist(unlist(variavel_list[i]), "triang", method="mge", start = list(min=-1, mode=0,max=1))
        parametro <- unlist(fit_tri['estimate'])
        min_var <- as.numeric(parametro[1])
        moda_var <- as.numeric(parametro[2])
        max_var <- as.numeric(parametro[3])
        dist.tri  <-   rtri(ITER$Qtd_iteracoes, min_var,  max_var, moda_var)
        aic_tri <- aic(unlist(variavel_list[i]), dist.tri)
     
     #FITTING Distribuicao PERT
        fit_pert <- fitdist(unlist(variavel_list[i]), "pert", method="mge",
                      start=list(min=-1, mode=0, max=1, shape=1),
                      lower=c(-Inf, -Inf, -Inf, 0), upper=c(Inf, Inf, Inf, Inf))
        parametro <-  unlist(fit_pert['estimate'])
        min_var   <-  as.numeric(parametro[1])
        max_var   <-  as.numeric(parametro[3])
        moda_var<-    as.numeric(parametro[2])
        shape_var<-   as.numeric(parametro[4])
        dist.pert <-  rpert(ITER$Qtd_iteracoes, min_var , moda_var, max_var, shape_var)
        aic_pert <-   aic(unlist(variavel_list[i]), dist.pert)
        
        
      #FITTING Distribuicao LOGNORMAL      
        fit_lnorm <- fitdist(unlist(variavel_list[i])+2, "lnorm", method="mge")
        parametro <- unlist(fit_lnorm['estimate'])
        media_var <- as.numeric(parametro[1])
        desv.p_var <- as.numeric(parametro[2])
        dist.lnorm <-   rlnormb(ITER$Qtd_iteracoes, media_var, desv.p_var)-2
        aic_lnorm <- aic(unlist(variavel_list[i]), dist.lnorm)

          
     matriz_dist <- data.frame('norm'=dist.norm,'tri'=dist.tri,'pert'=dist.pert,'lnorm'=dist.lnorm)
     
       
      teste_aic <- data.frame(
            'norm'=c('norm',aic_norm),
            'tri'=c('tri',aic_tri),
            'pert'=c('pert',aic_pert),
            'lnorm'=c('lnorm',aic_lnorm)
            )
     
      k1 <- which.min(teste_aic[2,])  # k1 se refere ao menor valor do teste de AIC nas distribuicoes testadas
     
      #Atribuir a melhor distribuição à cada variável
      assign(i,unlist(matriz_dist[teste_aic[1,k1]]))
      print(teste_aic)
      message('MELHOR DISTRIBUICAO PARA ',i,':',teste_aic[1,k1])
      

    
  }
  
      
#----------------- CALCULO DO CUSTO DE CAPITAL PROPRIO (Re) --------------- 

  if (lambda_desativado == FALSE) {
    Re_usd <- RF + BETA$`Beta`*(PRM)+BETA$`Lambda`*(RP)
  } else {
    Re_usd <- RF + BETA$`Beta`*(PRM)+1*(RP)
  }
  
  
  
  #Custo REAL do capital proprio
  
  Re_real <- (1+Re_usd)/(1+CPI)-1
  
 #spread do Custo de capital proprio do setor de ferrovias sobre o IdkA 5 anos
  
  Re_spread <-(Re_real-TLP_pre)

#-------------- CALCULO DO CUSTO DE CAPITAL DE TERCEIROS ------------ 

  
  ##custo de captacao (debentures e outros)
  custo_deb <- (((DEB_IPCA$Quant_IPCA/(DEB_IPCA$Quant_IPCA+DEB_DI$Quant_DI))*(DEB_IPCA$media_spread_IPCA+DEB_IPCA$media_cust_IPCA+IPCA+idka5))+((DEB_DI$Quant_DI/(DEB_IPCA$Quant_IPCA+DEB_DI$Quant_DI))*(DEB_DI$media_spread_DI+DEB_DI$media_cust_DI+CDI)))
  
  ##custo de financiamento (BNDES)
  custo_fin <- (1+TLP_pre)*(1+IPCA)*(1+CST_BNDES$`Remuneracao do BNDES - apoio direto`+CST_BNDES$`Taxa de risco de credito - apoio direto`)-1
  
  
  
  ##Custo Nominal do capital de terceiros (Rd)
  
  if (Rd_Bacen == FALSE) {
    Rd_nom <- (FONTES_FIN$`BNDES/`*custo_fin)+((FONTES_FIN$`Debentures/`+FONTES_FIN$`Outros/`)*custo_deb)
    Rd_nom2 <- (DIRE + TPB) / 2
    message(mean(Rd_nom))
  } else {
    Rd_nom <- (DIRE + TPB) / 2
    message(mean(Rd_nom))
  }
  
  
  
  ##Custo REAL do capital de terceiros  
  Rd_real <-(Rd_nom+1)/(IPCA+1)-1
  
  ##spread do Custo de Captacao de recursos de terceiros sobre o IdkA 5 anos
  Rd_spread <-Rd_real-TLP_pre
  
  #------------------------------ CMPC ------------------------------ 
  
  # calculo do CMPC real
  CMPC_real <- (EST_CAP$Proporcao_Equity*Re_real)+(EST_CAP$Proporcao_Debt*Rd_real*(1-EF_TRIBUT$Efeito_Tributario))
  

  # calculo do prêmio do CMPC sobre a TLP
  CMPC_spread <- CMPC_real-TLP_pre
  
  
  #--------------- PREPARACAO DOS RESULTADOS PARA O USUARIO  ----------------
  
  
  hist(CMPC_spread, breaks = 30) #plota o histograma de probabilidades obtidas para o CMPC
  
  message('media spread CMPC: ', round(mean(CMPC_spread)*10000,2),' bps')
  message('desvio padrao spread CMPC: ', round(sd(CMPC_spread)*10000,2),' bps')
  
  if (riscos_MF == FALSE) {
    CR0 <- -0.1
    CR1 <-  0.1
    CR2 <-  0.3
    CR3 <-  0.5
  } else {
    CR0 <-  0
    CR1 <-  0
    CR2 <-  0.5
    CR3 <-  1.0
  }
  
  
  
  message('CR 0: TLP-pre + ', round((mean(CMPC_spread)+CR0*sd(CMPC_spread))*10000,2),' bps')
  message('CR 1: TLP pre + ', round((mean(CMPC_spread)+CR1*sd(CMPC_spread))*10000,2),' bps')
  message('CR 2: TLP-pre + ', round((mean(CMPC_spread)+CR2*sd(CMPC_spread))*10000,2),' bps') 
  message('CR 3: TLP-pre + ', round((mean(CMPC_spread)+CR3*sd(CMPC_spread))*10000,2),' bps') 

  
  #---------------GRAVA DADOS DAS SIMULAÇÕES EM TXT--------------------
  
  linha_nova = paste(Anos,",", lambda_desativado,",", corrige_janela_PRM, ",",riscos_MF, ",", Rd_Bacen, ","  , round(mean(CMPC_spread)*10000,2),"," , round(sd(CMPC_spread)*10000,2), ",", round((mean(CMPC_spread)+CR0*sd(CMPC_spread))*10000,2), ",", round((mean(CMPC_spread)+CR1*sd(CMPC_spread))*10000,2), ",", round((mean(CMPC_spread)+CR2*sd(CMPC_spread))*10000,2), ",", round((mean(CMPC_spread)+CR3*sd(CMPC_spread))*10000,2), ","  , round(mean(Re_real)*10000,2), ","  , round(mean(Re_spread)*10000,2), sep = "")
  write(linha_nova,file="simu_adaptado_PERIODOS2.txt",append=TRUE)
  
  
  #---------------GRAVA DADOS DAS SIMULAÇÕES EM TXT--------------------
  
  
  # organização dos dados que serao inseridos na planilha "Resultado_CMPC"
  CMPC_distr <- data.frame(
                    Percentil = c(1:100),
                    CMPC_spread= quantile(CMPC_spread, seq(1/100, 1, 1/100))
                    )
  CMPC_distr[1,3] <- data.frame(media_CMPC_spread = mean(CMPC_spread))
  CMPC_distr[1,4] <- data.frame(desv.padrao_CMPC_spread = sd(CMPC_spread))
  
  
  
  # Impressão dos resultados na planilha "Resultado CMPC"
  #write_xlsx(CMPC_distr, paste(caminho, "/Resultado_CMPC", ".xlsx",sep=""))

  
  #==================================================================
  #
  #========================== FIM DO SCRIPT  ========================
  #
  #==================================================================
  
