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
                     "SN1", "SN2", "SN3", "SN4", "SN5",
                     "PR1_1", "PR1_2", "PR2_1", "PR2_2",
                     "PR3_1", "PR3_2", "PR4_1", "PR4_2",
                     "PR5_1", "PR5_2", "PR6_1", "PR6_2",
                     "TR1", "TR2", "TR3", "TR4", "TR5", "TR6", "TR7",
                     "UA1", "UA2", "UA3", "UA4", "UA5",
                     "PD1", "PD2", "PD3", "PD4", "PD5",
                     "IC1", "IC2", "IC3", "IC4", "IC5", "IC6")]

# Model Tanımlaması - Harici dosyadan oku
model <- readLines("model/measurement-model-v2.txt")
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

# Güvenilirlik analizi (Güncel fonksiyonlar)
compRelSEM(fit, tau.eq = FALSE)
AVE(fit)

# Model fit özeti
fitMeasures(fit, c("chisq", "df", "pvalue", "cfi", "tli",
                   "rmsea", "rmsea.ci.lower", "rmsea.ci.upper",
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

# 2. Parametre Tahminleri (Standardized ve Unstandardized)
std_solution <- standardizedSolution(fit)
param_estimates <- parameterEstimates(fit)

# Faktör yükleri (=~), korelasyonlar/kovaryanslar (~~), ve regresyonları (~) al
# Sadece gözlenen değişkenlerin rezidüel varyanslarını çıkar
# Latent faktör varyanslarını ve diğer tüm parametreleri tut
latent_vars <- lavNames(fit, type = "lv")
observed_vars <- lavNames(fit, type = "ov")

std_solution <- std_solution[std_solution$op %in% c("=~", "~~", "~"), ]
param_estimates <- param_estimates[param_estimates$op %in% c("=~", "~~", "~"), ]

# Sadece gözlenen değişkenlerin rezidüel varyanslarını çıkar (lhs == rhs VE lhs observed variable)
std_solution <- std_solution[!(std_solution$op == "~~" &
                                std_solution$lhs == std_solution$rhs &
                                std_solution$lhs %in% observed_vars), ]
param_estimates <- param_estimates[!(param_estimates$op == "~~" &
                                      param_estimates$lhs == param_estimates$rhs &
                                      param_estimates$lhs %in% observed_vars), ]

# Unstandardized estimates'i ekle
std_solution$est.unstd <- param_estimates$est

# Sütun sıralamasını ayarla (est.unstd, rhs ve est.std arasına)
col_order <- c("lhs", "op", "rhs", "est.unstd", "est.std", "se", "z",
               "pvalue", "ci.lower", "ci.upper")
std_solution <- std_solution[, col_order]

# ~~ operatörü için lhs boş olanları düzelt (rhs'i lhs'e taşı)
variance_rows <- std_solution$op == "~~" & std_solution$lhs == ""
if (any(variance_rows)) {
  std_solution$lhs[variance_rows] <- std_solution$rhs[variance_rows]
  std_solution$rhs[variance_rows] <- std_solution$lhs[variance_rows]
}

# Sayısal sütunları 3 ondalık basamağa yuvarla
numeric_cols <- c("est.unstd", "est.std", "se", "z", "pvalue", "ci.lower", "ci.upper")
std_solution[numeric_cols] <- lapply(std_solution[numeric_cols], function(x) round(x, 3))

# Faktör sırasını belirle
factor_order <- c("BI", "PU", "PEU", "SN", "TR", "UA", "PD", "IC",
                  "PR1", "PR2", "PR3", "PR4", "PR5", "PR6", "PR")

# lhs sütununda faktör sıralamasını uygula
std_solution$lhs <- factor(std_solution$lhs, levels = factor_order)
std_solution <- std_solution[order(std_solution$op, std_solution$lhs), ]
std_solution$lhs <- as.character(std_solution$lhs)

addWorksheet(wb, "Parameter_Estimates")
writeData(wb, "Parameter_Estimates", std_solution)

# 3. Güvenilirlik Analizi (Güncel fonksiyonlar)
comp_rel <- compRelSEM(fit, tau.eq = FALSE)  # Omega (composite reliability)
ave_results <- AVE(fit)  # Average Variance Extracted
alpha_results <- compRelSEM(fit, tau.eq = TRUE)  # Cronbach's Alpha (tau-equivalent)

# PR (second-order factor) için güvenilirlik hesapla
pr_subfactors <- c("PR1", "PR2", "PR3", "PR4", "PR5", "PR6")
rsq_all <- inspect(fit, "r2")
pr_rsq <- rsq_all[pr_subfactors]

# PR için AVE: alt faktörlerin R-squared ortalaması
ave_results["PR"] <- mean(pr_rsq, na.rm = TRUE)

# PR için Omega ve Alpha: alt faktörlerin ortalaması
comp_rel["PR"] <- mean(comp_rel[pr_subfactors], na.rm = TRUE)
alpha_results["PR"] <- mean(alpha_results[pr_subfactors], na.rm = TRUE)

# Güvenilirlik tablosu (Omega, Alpha, AVE)
rel_df <- data.frame(
  Construct = names(comp_rel),
  Omega = round(as.numeric(comp_rel), 3),
  Alpha = round(as.numeric(alpha_results[names(comp_rel)]), 3),
  AVE = round(as.numeric(ave_results[names(comp_rel)]), 3)
)
rownames(rel_df) <- NULL

# Faktör sırasını uygula
rel_df$Construct <- factor(rel_df$Construct, levels = factor_order)
rel_df <- rel_df[order(rel_df$Construct), ]
rel_df$Construct <- as.character(rel_df$Construct)

addWorksheet(wb, "Reliability")
writeData(wb, "Reliability", rel_df)

# 4. R-Square değerleri (varsa)
if (length(inspect(fit, "r2")) > 0) {
  rsq <- inspect(fit, "r2")
  rsq_df <- data.frame(
    Variable = names(rsq),
    R_Square = round(as.numeric(rsq), 3)
  )

  addWorksheet(wb, "R_Square")
  writeData(wb, "R_Square", rsq_df)
}

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

# 6. Fornell-Larcker Kriteri
# Ana faktörler (PR alt boyutları yerine üst düzey PR faktörü)
main_factors <- c("BI", "PU", "PEU", "SN", "TR", "UA", "PD", "IC", "PR")

# Korelasyon matrisini al
cor_matrix <- lavInspect(fit, "cor.lv")
cor_matrix <- cor_matrix[main_factors, main_factors]

# AVE değerlerinin karekökünü diagonal'e yerleştir
sqrt_ave <- sqrt(ave_results[main_factors])
diag(cor_matrix) <- sqrt_ave

# Fornell-Larcker matrisi
fornell_larcker <- as.data.frame(round(cor_matrix, 3))
fornell_larcker <- cbind(Construct = rownames(fornell_larcker), fornell_larcker)
rownames(fornell_larcker) <- NULL

addWorksheet(wb, "Fornell_Larcker")
writeData(wb, "Fornell_Larcker", fornell_larcker)

# 7. HTMT (Heterotrait-Monotrait Ratio)
cor_lv <- lavInspect(fit, "cor.lv")

# Ana faktörler için HTMT matrisi oluştur
n_factors <- length(main_factors)
htmt_matrix <- matrix(1, n_factors, n_factors)
rownames(htmt_matrix) <- main_factors
colnames(htmt_matrix) <- main_factors

# HTMT'yi korelasyon matrisinden hesapla
for (i in 1:(n_factors-1)) {
  for (j in (i+1):n_factors) {
    htmt_val <- abs(cor_lv[main_factors[i], main_factors[j]])
    htmt_matrix[i, j] <- htmt_val
    htmt_matrix[j, i] <- htmt_val
  }
}

htmt_df <- as.data.frame(round(htmt_matrix, 3))
htmt_df <- cbind(Construct = rownames(htmt_df), htmt_df)
rownames(htmt_df) <- NULL

addWorksheet(wb, "HTMT")
writeData(wb, "HTMT", htmt_df)

# 8. Info sayfası - Terim açıklamaları
info_df <- data.frame(
  Term = c("Omega", "α (Alpha)", "AVE", "=~", "~~", "~",
           "PR (Second-Order Factor)", "PR AVE", "PR Omega", "PR Alpha"),
  Description = c(
    "McDonald's Omega: Composite reliability coefficient based on factor loadings. Values > 0.70 indicate acceptable reliability.",
    "Cronbach's Alpha: Internal consistency reliability coefficient. Values > 0.70 indicate acceptable reliability. Less robust than Omega for multidimensional constructs.",
    "Average Variance Extracted: Proportion of variance explained by the construct relative to measurement error. Values > 0.50 indicate adequate convergent validity.",
    "Factor loading operator: Indicates measured variable is loaded by latent variable (e.g., 'BI =~ BI1' means BI1 is an indicator of BI).",
    "Variance/Covariance operator: Indicates variance (when lhs=rhs) or covariance/correlation (when lhs!=rhs) between variables.",
    "Regression operator: Indicates direct effect or regression path from rhs to lhs (e.g., 'BI ~ PU' means PU predicts BI).",
    "PR is a second-order factor composed of first-order subfactors (PR1-PR6). Reliability metrics for PR are calculated as averages of its subfactors.",
    "PR AVE: Average of R-squared values of subfactors (PR1-PR6). Represents the proportion of variance in subfactors explained by the second-order PR factor.",
    "PR Omega: Average of Omega values of subfactors (PR1-PR6). Represents the composite reliability of the second-order PR factor.",
    "PR Alpha: Average of Cronbach's Alpha values of subfactors (PR1-PR6). Represents the internal consistency of the second-order PR factor."
  )
)

addWorksheet(wb, "Info")
writeData(wb, "Info", info_df)

# Excel dosyasını kaydet
output_file <- "output/measurement-model-v2-results.xlsx"
dir.create("output", showWarnings = FALSE, recursive = TRUE)
saveWorkbook(wb, output_file, overwrite = TRUE)

cat("\n==============================================\n")
cat("Measurement Model sonuçları Excel'e yazıldı:\n")
cat(output_file, "\n")
cat("==============================================\n")
