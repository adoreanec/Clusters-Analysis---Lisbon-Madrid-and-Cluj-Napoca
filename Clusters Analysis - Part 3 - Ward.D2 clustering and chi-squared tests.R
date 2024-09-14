library(readxl)
library(dplyr)
library(cluster)
library(FactoMineR)
library(writexl)
library(purrr)
library(factoextra)
library(tidyr)
library(stringr)
library(ggplot2)

# Load data
df <- read_excel("C:/.../path_to_database.xlsx")

# Set the known parameters
cidades <- unique(df$cidd)
metodos <- list(
  "Lisboa" = list(method = "ward.D2", k = 5),
  "Cluj-Napoca" = list(method = "ward.D2", k = 5),
  "Madrid" = list(method = "ward.D2", k = 5)
)

# Definition of variables to be used
categorical_vars <- c("btp", "ner", "gen", "idd", "spr", "caf", "dvr", "daf_c",
                      "mtvr_walking", "mtvr_ownmicromob", "mtvr_sharedmicromob",
                      "mtvr_pt", "mtvr_car")

# Main loop
for (cidade in cidades) {
  method_info <- metodos[[cidade]]
  df_filtered <- df %>%
    filter(cidd == cidade) %>%
    mutate(across(categorical_vars, as.factor))
  
  if (nrow(df_filtered) > 0) {
    jaccard_distance <- daisy(df_filtered[categorical_vars], metric = "gower")
    hclust_result <- hclust(jaccard_distance, method = method_info$method)
    clusters <- cutree(hclust_result, k = method_info$k)
    
    # Add 'cluster' column
    df_filtered <- df_filtered %>%
      mutate(cluster = as.factor(clusters))
    
    # Calculation of the percentages of each categorisation in relation to the total of each category
    percentages <- lapply(categorical_vars, function(var) {
      cluster_percentages <- lapply(levels(df_filtered$cluster), function(cluster_level) {
        cluster_df <- subset(df_filtered, cluster == cluster_level)
        cluster_category_counts <- table(cluster_df[[var]])
        cluster_prop_table <- (cluster_category_counts / sum(cluster_category_counts)) * 100
        
        return(data.frame(cluster = cluster_level,
                          variable = var,
                          category = names(cluster_prop_table),
                          percentage = cluster_prop_table))
      })
      
      return(do.call(rbind, cluster_percentages))
    }) %>% bind_rows()
    
    # Chi-square tests
    chi_results <- lapply(categorical_vars, function(var) {
      test <- chisq.test(table(df_filtered$cluster, df_filtered[[var]]))
      return(data.frame(variable = var, Chi_Square = test$statistic, p_value = test$p.value))
    }) %>% bind_rows()
    
    # Export the results and the database with assigned clusters
    excel_filename <- paste0("Cluster_Results_For_", cidade, ".xlsx")
    write_xlsx(list(
      ClusterPercentages = percentages,
      ChiSquaredTests = chi_results,
      ClusteredData = df_filtered
    ), path = excel_filename)
    
  } else {
    message("Nenhum dado para ", cidade)
  }
}

