library(readxl)
library(dplyr)
library(cluster)
library(factoextra)
library(fpc)
library(klaR)
library(clValid)
library(writexl)

# Load data
dataset <- read_excel("C:/.../path_to_database.xlsx")
df <- data.frame(dataset)

# Definition of cities
cidades <- unique(df$cidd)

# Definition of variables to be used
categorical_vars <- c("btp", "ner", "gen", "idd", "spr", "caf", "dvr", "daf_c",
                      "mtvr_walking", "mtvr_ownmicromob", "mtvr_sharedmicromob",
                      "mtvr_pt", "mtvr_car", "mtvr_taxi")

# Clustering methods for testing
linkage_methods <- c("ward.D2", "single", "complete", "average")

# Loop to calculate dendrograms and export them
for (cidade in cidades) {
  df_filtered <- df %>%
    filter(cidd == cidade) %>%
    mutate(across(categorical_vars, as.factor))
  
  if (nrow(df_filtered) > 0) {
    for (method in linkage_methods) {
      # Calculate the Jaccard distance
      jaccard_distance <- daisy(df_filtered[categorical_vars], metric = "gower")
      
      # Hierarchical Clustering
      hclust_result <- hclust(jaccard_distance, method = method)
      
      # Export dendrograms
      dend <- fviz_dend(hclust_result, cex = 0.5)
      ggsave(paste0("Dendrogram_", cidade, "_", method, ".png"), dend)
    }
  } else {
    cat("Nenhum dado para", cidade, "\n")
  }
}

# Number of clusters to test
num_clusters <- 2:10

# Create dataframe to store all results
all_results <- data.frame()

for (cidade in cidades) {
  for (method in linkage_methods) {
    for (k in num_clusters) {
      df_filtered <- df %>%
        filter(cidd == cidade) %>%
        mutate(across(categorical_vars, as.factor))
      
      if (nrow(df_filtered) > 0) {
        jaccard_distance <- daisy(df_filtered[categorical_vars], metric = "gower")
        
        # Hierarchical Clustering
        hclust_result <- hclust(jaccard_distance, method = method)
        clusters <- cutree(hclust_result, k = k)
        
        # Clusters evaluation
        silhouette_score <- mean(silhouette(clusters, jaccard_distance)[, "sil_width"])
        dunn_score <- clValid::dunn(dist(as.matrix(jaccard_distance)), clusters)
        ch_score <- calinhara(jaccard_distance, clusters)
        
        # Save the results in the cumulative dataframe
        all_results <- rbind(all_results, data.frame(
          City = cidade,
          Method = method,
          NumClusters = k,
          Silhouette = silhouette_score,
          Dunn = dunn_score,
          CalinskiHarabasz = ch_score
        ))
      } else {
        cat("Nenhum dado para", cidade, " com o método ", method, " e ", k, " clusters.\n")
      }
    }
  }
}

# Export the accumulated results to a single Excel file
write_xlsx(all_results, path = "All_Cluster_Results.xlsx")
