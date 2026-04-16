# Gerekli kütüphaneleri yükle
library(openxlsx)
library(moments)  # skewness ve kurtosis için

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
analysis_vars <- c("BI1", "BI2", "BI3", "BI4", "BI5",
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

data <- data_all[, analysis_vars]

# Descriptive statistics hesapla
desc_stats <- data.frame(
  Variable = analysis_vars,
  Mean = round(sapply(data, mean, na.rm = TRUE), 2),
  SD = round(sapply(data, sd, na.rm = TRUE), 3),
  Skewness = round(sapply(data, skewness, na.rm = TRUE), 3),
  Kurtosis = round(sapply(data, kurtosis, na.rm = TRUE), 3)
)

# Satır isimlerini temizle
rownames(desc_stats) <- NULL

# Sonuçları görüntüle
print(desc_stats)

# Excel'e yaz
wb <- createWorkbook()
addWorksheet(wb, "Descriptive_Statistics")

# Ana tabloyu yaz
writeData(wb, "Descriptive_Statistics", desc_stats)

# Boş satır bırak ve Minimum/Maximum değerlerini ekle
start_row <- nrow(desc_stats) + 3  # +2 (header + data rows) + 1 (boş satır)

# Minimum değerleri hesapla (data frame olarak)
min_df <- data.frame(
  Variable = "Minimum",
  Mean = round(min(desc_stats$Mean, na.rm = TRUE), 2),
  SD = round(min(desc_stats$SD, na.rm = TRUE), 3),
  Skewness = round(min(desc_stats$Skewness, na.rm = TRUE), 3),
  Kurtosis = round(min(desc_stats$Kurtosis, na.rm = TRUE), 3)
)

# Maximum değerleri hesapla (data frame olarak)
max_df <- data.frame(
  Variable = "Maximum",
  Mean = round(max(desc_stats$Mean, na.rm = TRUE), 2),
  SD = round(max(desc_stats$SD, na.rm = TRUE), 3),
  Skewness = round(max(desc_stats$Skewness, na.rm = TRUE), 3),
  Kurtosis = round(max(desc_stats$Kurtosis, na.rm = TRUE), 3)
)

# Minimum satırını yaz (number formatında)
writeData(wb, "Descriptive_Statistics",
          min_df,
          startRow = start_row,
          colNames = FALSE)

# Maximum satırını yaz (number formatında)
writeData(wb, "Descriptive_Statistics",
          max_df,
          startRow = start_row + 1,
          colNames = FALSE)

# Excel dosyasını kaydet
output_file <- "output/descriptive-statistics-results.xlsx"
dir.create("output", showWarnings = FALSE, recursive = TRUE)
saveWorkbook(wb, output_file, overwrite = TRUE)

cat("\n==============================================\n")
cat("Descriptive statistics Excel'e yazıldı:\n")
cat(output_file, "\n")
cat("==============================================\n")
