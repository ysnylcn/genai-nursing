# Gerekli kütüphaneleri yükle
library(MVN)
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

# Analizde kullanılacak tüm değişkenleri seç (Model değişkenleri)
model_vars <- c("BI1", "BI2", "BI3", "BI4", "BI5",
                "PU1", "PU2", "PU3", "PU4", "PU5",
                "PEU1", "PEU2", "PEU3", "PEU4", "PEU5",
                "SN1", "SN2", "SN3", "SN4", "SN5",
                "PR1_1", "PR1_2", "PR2_1", "PR2_2",
                "PR3_1", "PR3_2", "PR4_1", "PR4_2",
                "PR5_1", "PR5_2", "PR6_1", "PR6_2",
                "TR1", "TR2", "TR3", "TR4", "TR5", "TR6", "TR7",
                "UA1", "UA2", "UA3", "UA4", "UA5",
                "PD1", "PD2", "PD3", "PD4", "PD5",
                "IC1", "IC2", "IC3", "IC4", "IC5", "IC6")

data_model <- data_all[, model_vars]

# Eksik değerleri çıkar (MVN paketi için)
data_complete <- na.omit(data_model)

cat("==============================================\n")
cat("Normality Analizi\n")
cat("==============================================\n")
cat("Toplam gözlem sayısı:", nrow(data_model), "\n")
cat("Eksik değer olmayan gözlem sayısı:", nrow(data_complete), "\n")
cat("Değişken sayısı:", ncol(data_complete), "\n")
cat("==============================================\n\n")

# ============================================================================
# Multivariate Normality Test - Mardia's Test
# ============================================================================

cat("Multivariate Normality Test (Mardia's Test) çalıştırılıyor...\n\n")

# Mardia testi (skewness ve kurtosis)
mvn_result <- mvn(data = data_complete,
                  mvn_test = "mardia",
                  univariate_test = "SW",
                  descriptives = TRUE)

# Multivariate normality sonuçları
# NOT: MVN paketinin MVN sütunu hatalı sonuç verebilir, p-value'dan kendimiz türetiyoruz
# p.value sütunu "<0.001" gibi string olabilir, bu yüzden as.numeric NA döndürür
parse_p <- function(p_str) {
  p_str <- trimws(p_str)
  if (grepl("^<", p_str)) return(0)  # "<0.001" gibi değerler kesinlikle < 0.05
  return(as.numeric(p_str))
}

mv_norm <- mvn_result$multivariate_normality
mv_p_parsed <- sapply(mv_norm$p.value, parse_p)
mv_norm$MVN <- ifelse(mv_p_parsed < 0.05, "NO", "YES")

cat("\n==============================================\n")
cat("MULTIVARIATE NORMALITY RESULTS (MARDIA'S TEST)\n")
cat("==============================================\n")
print(mv_norm)
cat("\n")
skew_stat <- as.numeric(mv_norm$Statistic[mv_norm$Test == "Mardia Skewness"])
skew_p <- mv_norm$p.value[mv_norm$Test == "Mardia Skewness"]
kurt_stat <- as.numeric(mv_norm$Statistic[mv_norm$Test == "Mardia Kurtosis"])
kurt_p <- mv_norm$p.value[mv_norm$Test == "Mardia Kurtosis"]

# Değişken sayısı (p) ve gözlem sayısı (n)
p <- ncol(data_complete)
n <- nrow(data_complete)

# Beklenen kurtosis değeri: p(p+2)
expected_kurtosis <- p * (p + 2)

cat("==============================================\n")
cat("DETAILED MARDIA'S TEST REPORT\n")
cat("==============================================\n")
cat(sprintf("Mardia's multivariate skewness: %.3f (p %s)\n",
            skew_stat, skew_p))
cat(sprintf("Mardia's multivariate kurtosis: %.3f (p %s)\n",
            kurt_stat, kurt_p))
cat(sprintf("\nNumber of variables (p): %d\n", p))
cat(sprintf("Sample size (n): %d\n", n))
cat(sprintf("Expected kurtosis under normality: %d [p(p+2)]\n", expected_kurtosis))
cat("==============================================\n\n")

# Univariate normality sonuçları
uv_norm_print <- mvn_result$univariate_normality
uv_p_parsed <- sapply(uv_norm_print$p.value, parse_p)
uv_norm_print$Normality <- ifelse(uv_p_parsed < 0.05, "NO", "YES")

cat("==============================================\n")
cat("UNIVARIATE NORMALITY RESULTS (Shapiro-Wilk)\n")
cat("==============================================\n")
print(uv_norm_print)
cat("\n")

# Descriptive statistics
cat("==============================================\n")
cat("DESCRIPTIVE STATISTICS\n")
cat("==============================================\n")
print(mvn_result$descriptives)
cat("\n")

# ============================================================================
# Excel'e Çıktıları Yaz
# ============================================================================

cat("Excel dosyası oluşturuluyor...\n")

# Excel workbook oluştur
wb <- createWorkbook()

# 1. Multivariate Normality (mv_norm zaten yukarıda düzeltilmiş MVN sütunuyla mevcut)
mv_norm_df <- data.frame(
  Test = mv_norm$Test,
  Statistic = round(as.numeric(mv_norm$Statistic), 3),
  p_value = mv_norm$p.value,
  Method = mv_norm$Method,
  Result = ifelse(sapply(mv_norm$p.value, parse_p) < 0.05, "Non-Normal", "Normal")
)

addWorksheet(wb, "Multivariate_Normality")
writeData(wb, "Multivariate_Normality", mv_norm_df)

# 2. Univariate Normality
uv_norm <- mvn_result$univariate_normality
uv_norm_df <- data.frame(
  Test = uv_norm$Test,
  Variable = uv_norm$Variable,
  Statistic = round(as.numeric(uv_norm$Statistic), 3),
  p_value = uv_norm$p.value,
  Result = ifelse(sapply(uv_norm$p.value, parse_p) < 0.05, "Non-Normal", "Normal")
)

addWorksheet(wb, "Univariate_Normality")
writeData(wb, "Univariate_Normality", uv_norm_df)

# 3. Descriptive Statistics
desc_stats <- mvn_result$descriptives
desc_numeric_cols <- c("Mean", "Std.Dev", "Median", "Min", "Max", "25th", "75th",
                      "Skew", "Kurtosis")
desc_stats[desc_numeric_cols] <- lapply(desc_stats[desc_numeric_cols], function(x) {
  round(as.numeric(x), 3)
})

addWorksheet(wb, "Descriptive_Statistics")
writeData(wb, "Descriptive_Statistics", desc_stats)

# 4. Özet bilgi sayfası
summary_info <- data.frame(
  Metric = c("Total Observations",
            "Complete Cases",
            "Number of Variables (p)",
            "Missing Data Rate (%)",
            "Mardia Skewness Statistic",
            "Mardia Skewness p-value",
            "Mardia Kurtosis Statistic",
            "Mardia Kurtosis p-value",
            "Expected Kurtosis [p(p+2)]",
            "Mardia Skewness Result",
            "Mardia Kurtosis Result"),
  Value = c(nrow(data_model),
           nrow(data_complete),
           ncol(data_complete),
           round((1 - nrow(data_complete)/nrow(data_model)) * 100, 2),
           round(skew_stat, 3),
           skew_p,
           round(kurt_stat, 3),
           kurt_p,
           expected_kurtosis,
           mv_norm_df$Result[mv_norm_df$Test == "Mardia Skewness"],
           mv_norm_df$Result[mv_norm_df$Test == "Mardia Kurtosis"])
)

addWorksheet(wb, "Summary")
writeData(wb, "Summary", summary_info)

# Excel dosyasını kaydet
output_file <- "output/normality-analysis-results.xlsx"
dir.create("output", showWarnings = FALSE, recursive = TRUE)
saveWorkbook(wb, output_file, overwrite = TRUE)

cat("\n==============================================\n")
cat("Normality analizi sonuçları Excel'e yazıldı:\n")
cat(output_file, "\n")
cat("==============================================\n")

# Yorum ve Öneriler
cat("\n==============================================\n")
cat("INTERPRETATION NOTES\n")
cat("==============================================\n")
cat("Mardia's Kurtosis Test:\n")
cat("  - If p < 0.05: Multivariate non-normality detected\n")
cat("  - Recommendation: Use robust estimators (MLR, MLM) in SEM\n")
cat("\nShapiro-Wilk Test (Univariate):\n")
cat("  - If p < 0.05: Univariate non-normality for that variable\n")
cat("  - Large samples (N > 200) may show significant results\n")
cat("    even with minor deviations from normality\n")
cat("\nSkewness and Kurtosis Guidelines:\n")
cat("  - Skewness: Values between -1 and +1 are acceptable\n")
cat("  - Kurtosis: Values between -1 and +1 are excellent\n")
cat("            Values between -2 and +2 are acceptable\n")
cat("==============================================\n")
