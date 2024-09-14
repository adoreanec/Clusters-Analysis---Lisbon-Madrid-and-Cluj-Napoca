library(FactoMineR)
library(factoextra)
library(readxl)
library(dplyr)
library(writexl)

# Load data
dataset <- read_excel("C:/.../path_to_database.xlsx")
df <- data.frame(dataset)

# Select specific variables and convert to factors
vars_to_select <- c("btp", "ner", "gen", "idd", "sf", "imc", "spr", "freg",
                    "dmaf", "dspbt", "ccaf", "caf", "nvd", "dvr",
                    "dpvr", "mvr", "nvd", "daf_c", "daf_par_dis",
                    "dr", "mtvr_walking", "mtvr_ownmicromob", "mtvr_sharedmicromob",
                    "mtvr_pt", "mtvr_taxi", "mtvr_car")
df_selected <- df %>%
  select(all_of(vars_to_select)) %>%
  mutate(across(everything(), as.factor))

# Perform MCA (Multiple Correspondence Analysis)
mca_results <- MCA(df_selected, graph = FALSE)

# Calculate the total contribution of each variable
contributions <- as.data.frame(mca_results$var$contrib)
contributions$Variable <- rep(names(df_selected), sapply(df_selected, nlevels))
total_contributions <- aggregate(. ~ Variable, data = contributions, FUN = sum)

# Export total contributions to Excel
write_xlsx(list(Total_Contributions = total_contributions), "MCA_results.xlsx")

# Optional visualisation in R with fviz_mca_biplot
fviz_mca_biplot(mca_results, choice = "var", label = "var", repel = TRUE,
                gradient.cols = c("#00AFBB", "#E7B800", "#FC4E07"))

# Obtain the data for the MCA variables
mca_vars <- get_mca_var(mca_results)

# Ensure that ‘quality_representation’ is a data frame
quality_representation <- as.data.frame(mca_vars$cos2)

# Add the variable names as a new column
quality_representation$Variable <- rownames(quality_representation)

# Extract the name of the global variable
quality_representation$GlobalVariable <- sub("_.*", "", quality_representation$Variable)

# Calculate the total cos2 per variable
quality_representation$Total_cos2 <- rowSums(quality_representation[, sapply(quality_representation, is.numeric)])

# Calculate the average cos2 per global variable
global_cos2 <- quality_representation %>%
  group_by(GlobalVariable) %>%
  summarise(Average_cos2 = mean(Total_cos2))

# Organise variables by quality of average representation
ordered_quality <- global_cos2 %>%
  arrange(desc(Average_cos2))

# Export the results of the average cos2 per global variable to Excel
write_xlsx(list(Quality_Representation = ordered_quality), "MCA_cos2_results.xlsx")

