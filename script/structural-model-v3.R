# Gerekli kütüphaneleri yükle
library(lavaan)
library(semPlot)
library(semTools)
library(openxlsx)

getwd()

# Veriyi oku - Tüm değişkenler
data_all <- read.table("data/mplus.dat",
                       header = FALSE,
                       sep = "",
                       col.names = c("BI1", "BI2", "BI3", "BI4", "BI5",
                                    "PU1", "PU2", "PU3", "PU4", "PU5",
                                    "PEU1", "PEU2", "PEU3", "PEU4", "PEU5",
                                    "SN1", "SN2", "SN3", "SN4", "SN5",
                                    "PR1_1", "PR1_2", "PR2_1", "PR2_2",
                                    "PR3_1", "PR3_2", "PR4_1", "PR4_2",
                                    "PR5_1", "PR5_2", "PR6_1", "PR6_2",
                                    "TR1", "TR2", "TR3", "TR4", "TR5", "TR6", "TR7",
                                    "UA1", "UA2", "UA3", "UA4", "UA5",
                                    "PD1", "PD2", "PD3", "PD4", "PD5",
                                    "IC1", "IC2", "IC3", "IC4", "IC5", "IC6"),
                       na.strings = "-999")

# Analizde kullanılacak değişkenleri seç
data <- data_all[, c("BI1", "BI2", "BI3", "BI4", "BI5",
                     "PU1", "PU2", "PU3", "PU4", "PU5",
                     "PEU1", "PEU2", "PEU3", "PEU4", "PEU5",
                     "SN1", "SN2", "SN3", "SN5",
                     "PR1_1", "PR1_2", "PR2_1", "PR2_2",
                     "PR3_1", "PR3_2", "PR4_1", "PR4_2",
                     "PR5_1", "PR5_2", "PR6_1", "PR6_2",
                     "TR1", "TR2", "TR3", "TR4", "TR5", "TR6", "TR7",
                     "UA1", "UA2", "UA3", "UA4", "UA5",
                     "PD1", "PD2", "PD3", "PD4", "PD5",
                     "IC3", "IC4", "IC5", "IC6")]

# Model Tanımlaması - Harici dosyadan oku
model <- readLines("model/structural-model-v3.txt")
model <- paste(model, collapse = "\n")

# Modeli çalıştır
# NOT: MLM kullanılıyor çünkü normality analizi multivariate non-normality gösterdi
fit <- sem(model,
           data = data,
           estimator = "MLM",      # Robust ML (Satorra-Bentler) for non-normal data
           missing = "listwise",   # Listwise deletion (MLM ile FIML kullanılamaz)
           std.lv = FALSE)         # Lavaan default: first loading = 1

# Sonuçları görüntüle
summary(fit,
        fit.measures = TRUE,
        standardized = TRUE,
        rsquare = TRUE,
        ci = TRUE)

# Modification indices (>10)
modindices(fit, minimum.value = 10, sort = TRUE)

# Residuals
resid(fit, type = "standardized")

# Model fit özeti
fitMeasures(fit, c("chisq", "df", "pvalue",
                   "chisq.scaled", "pvalue.scaled",
                   "cfi", "cfi.robust",
                   "tli", "tli.robust",
                   "rmsea", "rmsea.robust",
                   "srmr", "aic", "bic"))

# Parametre tahminleri (standardize edilmiş)
standardizedSolution(fit)

# ============================================================================
# Excel'e Çıktıları Yaz
# ============================================================================

# Excel workbook oluştur
wb <- createWorkbook()

# 1. Model Fit İstatistikleri
# MLM kullanıldığında hem normal hem robust fit indices alınır
fit_stats <- fitMeasures(fit, c("chisq", "df", "pvalue",
                                "chisq.scaled", "pvalue.scaled",  # Satorra-Bentler
                                "cfi", "cfi.robust",
                                "tli", "tli.robust",
                                "rmsea", "rmsea.robust",
                                "rmsea.ci.lower", "rmsea.ci.upper",
                                "rmsea.ci.lower.robust", "rmsea.ci.upper.robust",
                                "srmr", "aic", "bic"))
fit_df <- data.frame(
  Measure = names(fit_stats),
  Value = round(as.numeric(fit_stats), 3)
)

addWorksheet(wb, "Model_Fit")
writeData(wb, "Model_Fit", fit_df)

# 2. Parametre Tahminleri (Standardize Edilmiş)
std_solution <- standardizedSolution(fit)
param_estimates <- parameterEstimates(fit)

# Sadece faktör yükleri (=~), korelasyonlar/kovaryanslar (~~), ve regresyonları (~) al
std_solution <- std_solution[std_solution$op %in% c("=~", "~~", "~"), ]
param_estimates <- param_estimates[param_estimates$op %in% c("=~", "~~", "~"), ]

# Sadece gözlenen değişkenlerin rezidüel varyanslarını çıkar
observed_vars <- lavNames(fit, type = "ov")
std_solution <- std_solution[!(std_solution$op == "~~" &
                                std_solution$lhs == std_solution$rhs &
                                std_solution$lhs %in% observed_vars), ]
param_estimates <- param_estimates[!(param_estimates$op == "~~" &
                                      param_estimates$lhs == param_estimates$rhs &
                                      param_estimates$lhs %in% observed_vars), ]

# Unstandardized estimates'i ekle
std_solution$est.unstd <- param_estimates$est

# Sütun sıralamasını ayarla
col_order <- c("lhs", "op", "rhs", "est.unstd", "est.std", "se", "z",
               "pvalue", "ci.lower", "ci.upper")
std_solution <- std_solution[, col_order]

# Sayısal sütunları 3 ondalık basamağa yuvarla
numeric_cols <- c("est.unstd", "est.std", "se", "z", "pvalue", "ci.lower", "ci.upper")
std_solution[numeric_cols] <- lapply(std_solution[numeric_cols], function(x) round(x, 3))

addWorksheet(wb, "Standardized_Solution")
writeData(wb, "Standardized_Solution", std_solution)

# 3. Path Coefficients (Sadece yapısal ilişkiler: ~)
path_coef <- std_solution[std_solution$op == "~", ]

# Hipotez etiketleri ekle
path_coef$Hypothesis <- NA
path_coef$Hypothesis[path_coef$lhs == "PU" & path_coef$rhs == "PEU"] <- "H1"
path_coef$Hypothesis[path_coef$lhs == "BI" & path_coef$rhs == "PU"]  <- "H2"
path_coef$Hypothesis[path_coef$lhs == "BI" & path_coef$rhs == "PEU"] <- "H3"
path_coef$Hypothesis[path_coef$lhs == "PU" & path_coef$rhs == "SN"]  <- "H4"
path_coef$Hypothesis[path_coef$lhs == "BI" & path_coef$rhs == "SN"]  <- "H5"
path_coef$Hypothesis[path_coef$lhs == "PU" & path_coef$rhs == "TR"]  <- "H6"
path_coef$Hypothesis[path_coef$lhs == "BI" & path_coef$rhs == "TR"]  <- "H7"
path_coef$Hypothesis[path_coef$lhs == "TR" & path_coef$rhs == "PR"]  <- "H8"
path_coef$Hypothesis[path_coef$lhs == "PU" & path_coef$rhs == "PR"]  <- "H9"
path_coef$Hypothesis[path_coef$lhs == "BI" & path_coef$rhs == "PR"]  <- "H10"
path_coef$Hypothesis[path_coef$lhs == "PR" & path_coef$rhs == "UA"]  <- "H11"
path_coef$Hypothesis[path_coef$lhs == "TR" & path_coef$rhs == "UA"]  <- "H12"
path_coef$Hypothesis[path_coef$lhs == "SN" & path_coef$rhs == "PD"]  <- "H13"
path_coef$Hypothesis[path_coef$lhs == "SN" & path_coef$rhs == "IC"]  <- "H14"

# Significance sütunu ekle
path_coef$Significance <- ifelse(path_coef$pvalue < 0.001, "***",
                           ifelse(path_coef$pvalue < 0.01, "**",
                             ifelse(path_coef$pvalue < 0.05, "*", "ns")))

# Hipotez sırasına göre sırala
path_coef <- path_coef[order(as.numeric(gsub("H", "", path_coef$Hypothesis))), ]

addWorksheet(wb, "Path_Coefficients")
writeData(wb, "Path_Coefficients", path_coef)

# 4. R-Square değerleri (Endojen değişkenler için)
rsq <- inspect(fit, "r2")

# Tüm endojen değişkenler için R²
rsq_all <- data.frame(
  Variable = names(rsq),
  R_Square = round(as.numeric(rsq), 3),
  Variance_Explained_Percent = paste0(round(as.numeric(rsq) * 100, 1), "%")
)

# Anahtar endojen yapılar için ayrı özet tablo
key_constructs <- c("BI", "PU", "TR", "PR", "SN")
rsq_key <- rsq[key_constructs]
rsq_key_df <- data.frame(
  Construct = names(rsq_key),
  R_Square = round(as.numeric(rsq_key), 3),
  Variance_Explained_Percent = paste0(round(as.numeric(rsq_key) * 100, 1), "%")
)

addWorksheet(wb, "R_Square_All")
writeData(wb, "R_Square_All", rsq_all)

addWorksheet(wb, "R_Square_Key_Constructs")
writeData(wb, "R_Square_Key_Constructs", rsq_key_df)

# 5. Modification Indices (>10)
mod_ind <- modindices(fit, minimum.value = 10, sort. = TRUE)

if (nrow(mod_ind) > 0) {
  # Sayısal sütunları 3 ondalık basamağa yuvarla
  mi_numeric_cols <- c("mi", "epc", "sepc.lv", "sepc.all", "sepc.nox")
  mod_ind[mi_numeric_cols] <- lapply(mod_ind[mi_numeric_cols], function(x) {
    if (is.numeric(x)) round(x, 3) else x
  })

  addWorksheet(wb, "Modification_Indices")
  writeData(wb, "Modification_Indices", mod_ind)
}

# 6. Hipotez Özet Tablosu
hyp_summary <- data.frame(
  Hypothesis = paste0("H", 1:14),
  Path = c("PEU -> PU", "PU -> BI", "PEU -> BI", "SN -> PU", "SN -> BI",
           "TR -> PU", "TR -> BI", "PR -> TR", "PR -> PU", "PR -> BI",
           "UA -> PR", "UA -> TR", "PD -> SN", "IC -> SN"),
  Theory = c(rep("TAM", 5), rep("Trust-Risk", 5), rep("Hofstede", 4))
)

# Path coefficients'tan beta ve p değerlerini eşleştir
path_lookup <- std_solution[std_solution$op == "~", ]
hyp_summary$Beta <- NA
hyp_summary$p_value <- NA
hyp_summary$Significance <- NA

# Her path için c(lhs, rhs) = c(outcome, predictor)
paths <- list(
  c("PU", "PEU"), c("BI", "PU"), c("BI", "PEU"), c("PU", "SN"), c("BI", "SN"),
  c("PU", "TR"), c("BI", "TR"), c("TR", "PR"), c("PU", "PR"), c("BI", "PR"),
  c("PR", "UA"), c("TR", "UA"), c("SN", "PD"), c("SN", "IC")
)

for (i in 1:14) {
  row <- path_lookup[path_lookup$lhs == paths[[i]][1] & path_lookup$rhs == paths[[i]][2], ]
  if (nrow(row) > 0) {
    hyp_summary$Beta[i] <- round(row$est.std, 3)
    hyp_summary$p_value[i] <- round(row$pvalue, 3)
    hyp_summary$Significance[i] <- ifelse(row$pvalue < 0.001, "***",
                                    ifelse(row$pvalue < 0.01, "**",
                                      ifelse(row$pvalue < 0.05, "*", "ns")))
  }
}

hyp_summary$Supported <- ifelse(hyp_summary$Significance != "ns", "Yes", "No")

addWorksheet(wb, "Hypothesis_Summary")
writeData(wb, "Hypothesis_Summary", hyp_summary)

# Excel dosyasını kaydet
output_file <- "output/structural-model-v3-results.xlsx"
dir.create("output", showWarnings = FALSE, recursive = TRUE)
saveWorkbook(wb, output_file, overwrite = TRUE)

cat("\n==============================================\n")
cat("Structural Model sonuçları Excel'e yazıldı:\n")
cat(output_file, "\n")
cat("==============================================\n")
