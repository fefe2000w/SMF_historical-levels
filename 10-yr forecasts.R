#### Before importing data, the missing values (NA) should be clear in excel file.
#### Otherwise the variables will be identified as categorical.


## Read historical data
setwd("D:/A Sem1 2024/FINM 6010")  # Change to your path

#install.packages("readxl")  # Make sure the package is installed
library(readxl)

sheet_names <- c("Economic drivers", "AU Equity", "DM Equity", "EM Equity")
dfs <- list()
for (sheet_name in sheet_names) {
  df <- read_excel("Historical data.xlsx", sheet=sheet_name)
  dfs[[sheet_name]] <- df
}

States <- c("HH", "HM", "HL", "MH", "MM", "ML", "LH", "LM", "LL", "Stagflation", "Crisis")

## Historical probabilities of each scenario state
probs <- function(x) {
  econ <- dfs[[x]]
  econ <- econ[complete.cases(econ), ]
  inf <- econ[["Inflation"]]
  gdp <- econ[["GDP / Potential"]]
  HH = HM = HL = MH = MM = ML = LH = LM = LL = Stag = Cris = 0
  for (i in 1:nrow(econ)) {
    if (inf[i] >= 7) {
      if (gdp[i]>=0.95 & gdp[i]<0.98) {
        Stag = Stag + 1
      }
    } else if (inf[i]>=4.5 & inf[i]<7) {
      if (gdp[i] >= 1.015) {
        HH = HH + 1
      } else if (gdp[i]>=0.995 & gdp[i]<1.015) {
        HM = HM + 1
      } else if (gdp[i]>=0.98 & gdp[i]<0.995) {
        HL = HL + 1
      }
    } else if (inf[i]>=2.5 & inf[i]<4.5) {
      if (gdp[i] >= 1.015) {
        MH = MH + 1
      } else if (gdp[i]>=0.995 & gdp[i]<1.015) {
        MM = MM + 1
      } else if (gdp[i]>=0.98 & gdp[i]<0.995) {
        ML = ML + 1
      }
    } else if (inf[i]>=1 & inf[i]<2.5) {
      if (gdp[i] >= 1.015) {
        LH = LH + 1
      } else if (gdp[i]>=0.995 & gdp[i]<1.015) {
        LM = LM + 1
      } else if (gdp[i]>=0.98 & gdp[i]<0.995) {
        LL = LL + 1
      }
    } else if (inf[i]>=0 & inf[i]<1) {
      if (gdp[i]>=0.935 & gdp[i]<0.95) {
        Cris = Cris + 1
      }
    }
  }
  n = HH+HM+HL+MH+MM+ML+LH+LM+LL+Stag+Cris
  print(n)
#  probs <- c(HH, HM, HL, MH, MM, ML, LH, LM, LL, Stag, Cris)
  probs <- c(HH*100/n, HM*100/n, HL*100/n, MH*100/n, MM*100/n, ML*100/n, LH*100/n, LM*100/n, LL*100/n, Stag*100/n, Cris*100/n)
  return(probs)
}
His_probs<- data.frame(States, probs("Economic drivers"))


## Historical mean levels of PE and ROE for each asset classes
scenario_pe <- function(x, y) {
  data <- merge(dfs[[x]], dfs[[y]], by="Period")
  columns <- c("Inflation", "GDP / Potential", "12mth Fwd PE Ratio")
  data <- data[complete.cases(data[, columns]), ]
  inf <- data[["Inflation"]]
  gdp <- data[["GDP / Potential"]]
  pe <- data[["12mth Fwd PE Ratio"]]
  HH = HH_pe = HM = HM_pe = HL = HL_pe = 0
  MH = MH_pe = MM = MM_pe = ML = ML_pe = 0
  LH = LH_pe = LM = LM_pe = LL = LL_pe = 0
  Stag = Stag_pe = Cris = Cris_pe = 0
  for (i in 1:nrow(data)) {
    if (inf[i] >= 7) {
      if (gdp[i]>=0.95 & gdp[i]<0.98) {
        Stag = Stag + 1
        Stag_pe = Stag_pe + pe[i]
        }
    } else if (inf[i]>=4.5 & inf[i]<7) {
      if (gdp[i] >= 1.015) {
        HH = HH + 1
        HH_pe = HH_pe + pe[i]
      } else if (gdp[i]>=0.995 & gdp[i]<1.015) {
        HM = HM + 1
        HM_pe = HM_pe + pe[i]
      } else if (gdp[i]>=0.98 & gdp[i]<0.995) {
        HL = HL + 1
        HL_pe = HL_pe + pe[i]
      }
    } else if (inf[i]>=2.5 & inf[i]<4.5) {
      if (gdp[i] >= 1.015) {
        MH = MH + 1
        MH_pe = MH_pe + pe[i]
      } else if (gdp[i]>=0.995 & gdp[i]<1.015) {
        MM = MM + 1
        MM_pe = MM_pe + pe[i]
      } else if (gdp[i]>=0.98 & gdp[i]<0.995) {
        ML = ML + 1
        ML_pe = ML_pe + pe[i]
      }
    } else if (inf[i]>=1 & inf[i]<2.5) {
      if (gdp[i] >= 1.015) {
        LH = LH + 1
        LH_pe = LH_pe + pe[i]
      } else if (gdp[i]>=0.995 & gdp[i]<1.015) {
        LM = LM + 1
        LM_pe = LM_pe + pe[i]
      } else if (gdp[i]>=0.98 & gdp[i]<0.995) {
        LL = LL + 1
        LL_pe = LL_pe + pe[i]
      }
    } else if (inf[i]>=0 & inf[i]<1) {
      if (gdp[i]>=0.935 & gdp[i]<0.95) {
        Cris = Cris + 1
        Cris_pe = Cris_pe + pe[i]
      }
    }
  }
  pe_mean <- c(HH_pe/HH, HM_pe/HM, HL_pe/HL, MH_pe/MH, MM_pe/MM, ML_pe/ML, LH_pe/LH, LM_pe/LM, LL_pe/LL, Stag_pe/Stag, Cris_pe/Cris)
  return(pe_mean)
}


scenario_roe <- function(x, y) {
  data <- merge(dfs[[x]], dfs[[y]], by="Period")
  columns <- c("Inflation", "GDP / Potential", "ROE")
  data <- data[complete.cases(data[, columns]), ]
  inf <- data[["Inflation"]]
  gdp <- data[["GDP / Potential"]]
  roe <- data[["ROE"]]
  HH = HH_roe = HM = HM_roe = HL = HL_roe = 0
  MH = MH_roe = MM = MM_roe = ML = ML_roe = 0
  LH = LH_roe = LM = LM_roe = LL = LL_roe = 0
  Stag = Stag_roe = Cris = Cris_roe = 0
  for (i in 1:nrow(data)) {
    if (inf[i] >= 7) {
      if (gdp[i]>=0.95 & gdp[i]<0.98) {
        Stag = Stag + 1
        Stag_roe = Stag_roe + roe[i]
      }
    } else if (inf[i]>=4.5 & inf[i]<7) {
      if (gdp[i] >= 1.015) {
        HH = HH + 1
        HH_roe = HH_roe + roe[i]
      } else if (gdp[i]>=0.995 & gdp[i]<1.015) {
        HM = HM + 1
        HM_roe = HM_roe + roe[i]
      } else if (gdp[i]>=0.98 & gdp[i]<0.995) {
        HL = HL + 1
        HL_roe = HL_roe + roe[i]
      }
    } else if (inf[i]>=2.5 & inf[i]<4.5) {
      if (gdp[i] >= 1.015) {
        MH = MH + 1
        MH_roe = MH_roe + roe[i]
      } else if (gdp[i]>=0.995 & gdp[i]<1.015) {
        MM = MM + 1
        MM_roe = MM_roe + roe[i]
      } else if (gdp[i]>=0.98 & gdp[i]<0.995) {
        ML = ML + 1
        ML_roe = ML_roe + roe[i]
      }
    } else if (inf[i]>=1 & inf[i]<2.5) {
      if (gdp[i] >= 1.015) {
        LH = LH + 1
        LH_roe = LH_roe + roe[i]
      } else if (gdp[i]>=0.995 & gdp[i]<1.015) {
        LM = LM + 1
        LM_roe = LM_roe + roe[i]
      } else if (gdp[i]>=0.98 & gdp[i]<0.995) {
        LL = LL + 1
        LL_roe = LL_roe + roe[i]
      }
    } else if (inf[i]>=0 & inf[i]<1) {
      if (gdp[i]>=0.935 & gdp[i]<0.95) {
        Cris = Cris + 1
        Cris_roe = Cris_roe + roe[i]
      }
    }
  }
  roe_mean <- c(HH_roe/HH, HM_roe/HM, HL_roe/HL, MH_roe/MH, MM_roe/MM, ML_roe/ML, LH_roe/LH, LM_roe/LM, LL_roe/LL, Stag_roe/Stag, Cris_roe/Cris)
  return(roe_mean)
}

AE_pe <- scenario_pe("Economic drivers", "AU Equity")
AE_roe <- scenario_roe("Economic drivers", "AU Equity")
DM_pe <- scenario_pe("Economic drivers", "DM Equity")
DM_roe <- scenario_roe("Economic drivers", "DM Equity")
EM_pe <- scenario_pe("Economic drivers", "EM Equity")
EM_roe <- scenario_roe("Economic drivers", "EM Equity")
Forecasts <- data.frame(States, AE_pe, AE_roe, DM_pe, DM_roe, EM_pe, EM_roe)