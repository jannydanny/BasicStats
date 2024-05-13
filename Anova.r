# Load necessary libraries
library(readxl)   # for reading Excel files
library(car)      # for ANOVA
library(dplyr)    # for data manipulation
library(agricolae) # for Tukey's test

# Step 1: Read the data from Excel
raw_data <- read_excel("excel file location")

# Step 2: Sanitize column names to avoid issues in formulas
names(raw_data) <- make.names(names(raw_data), unique = TRUE)

# Convert columns to numeric, handling non-numeric values
clean_data <- raw_data %>%
  mutate(across(.cols = where(is.character), .fns = ~as.numeric(as.character(.))))

# Ensure the 'Group' column is treated as a factor
clean_data$Group <- as.factor(clean_data$Group)

# Step 3: Perform ANOVA for each variable excluding the Group column
anova_results <- list()  # List to store results of ANOVAs
tukey_results <- list()  # List to store Tukey HSD results
group_column <- "Group"  # Assuming 'Group' is the name of your group identifier column

# Get all column names except the group column
variables <- setdiff(names(clean_data), group_column)

for (var in variables) {
  # Exclude rows with NA values in the current variable
  data_subset <- clean_data[!is.na(clean_data[[var]]), ]
  
  # Check if there are at least two groups with multiple data points
  group_counts <- table(data_subset$Group)
  sufficient_groups <- names(group_counts[group_counts > 1])
  
  if (length(sufficient_groups) >= 2) {
    # Filter data for sufficient groups only
    sufficient_data_subset <- data_subset[data_subset$Group %in% sufficient_groups, ]
    
    # Perform ANOVA
    formula <- as.formula(paste(var, "~", group_column))
    model <- aov(formula, data = sufficient_data_subset)
    summary_result <- summary(model)
    
    # Store the ANOVA result
    anova_results[[var]] <- summary_result
    
    # Check if model is significant before running Tukey's test
    if (summary_result[[1]]$'Pr(>F)'[1] < 0.05) {
      # Perform Tukey's HSD test
      tukey_result <- try(HSD.test(model, "Group", group = TRUE), silent = TRUE)
      if (!inherits(tukey_result, "try-error")) {
        tukey_results[[var]] <- tukey_result
      }
    }
  }
}

# Step 4: Output only significant ANOVA and corresponding Tukey's HSD results
significant_anova_vars <- names(anova_results)[sapply(anova_results, function(res) {
  is.list(res) && !is.null(res[[1]]) && res[[1]]$'Pr(>F)'[1] < 0.05
})]

for (var in significant_anova_vars) {
  cat("ANOVA Results for", var, ":\n")
  print(anova_results[[var]])
  cat("\n\n")
  
  if (var %in% names(tukey_results)) {
    cat("Tukey HSD Results for", var, ":\n")
    print(tukey_results[[var]]$group)
    cat("\n\n")
  }
}
