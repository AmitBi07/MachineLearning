# Install and load gbm
if (!requireNamespace("gbm", quietly = TRUE)) {
  install.packages("gbm")
  library(gbm)
}
library(gbm)
library(openxlsx)
library(caret)
library(ggplot2)
library(gridExtra)

# Load the data
df <- read.xlsx("E:/AMIT/Thesis Work/Data analysis/Wheat_diseaseSeverity.xlsx", sheet = "WheatRust_ana")
df <- as.data.frame(df)
head(df, 2)
str(df)

# Set seed for reproducibility
set.seed(100)
setwd("E:/AMIT/Thesis Work/Data analysis")
# Create training and test data (70:30 split)
index <- createDataPartition(df$DS, p = 0.7, list = FALSE, times = 1)
train_df <- df[index, ]
test_df <- df[-index, ]

# Remove NA rows from test_df
test_df <- na.omit(test_df)

# Standardize the data
preProcValues <- preProcess(train_df, method = c("center", "scale"))
train_df <- predict(preProcValues, train_df)
test_df <- predict(preProcValues, test_df)

# Train GBM model
gbm_model <- gbm(
  formula = DS ~ .,
  distribution = "gaussian",
  data = train_df,
  n.trees = 1000,
  interaction.depth = 3,
  shrinkage = 0.01,
  cv.folds = 5,
  n.cores = NULL, # Use all cores
  verbose = FALSE
)

# Summary of the model
summary(gbm_model)

# Best number of trees based on cross-validation
best_trees <- gbm.perf(gbm_model, method = "cv")

# Make predictions
gbm_train_pred <- predict(gbm_model, train_df, n.trees = best_trees)
gbm_test_pred <- predict(gbm_model, test_df, n.trees = best_trees)

# Metrics for training set
data_cal_gbm <- data.frame(Predicted = gbm_train_pred, Observed = train_df$DS)
r_squared_cal_gbm <- cor(data_cal_gbm$Observed, data_cal_gbm$Predicted)^2
RMSEC_gbm <- RMSE(pred = data_cal_gbm$Predicted, obs = data_cal_gbm$Observed)
nRMSEC_gbm <- RMSEC_gbm * 100 / (max(data_cal_gbm$Observed) - min(data_cal_gbm$Observed))
MAEC_gbm <- MAE(pred = data_cal_gbm$Predicted, obs = data_cal_gbm$Observed)

# Metrics for test set
data_val_gbm <- data.frame(Predicted = gbm_test_pred, Observed = test_df$DS)
r_squared_val_gbm <- cor(data_val_gbm$Observed, data_val_gbm$Predicted)^2
RMSEV_gbm <- RMSE(pred = data_val_gbm$Predicted, obs = data_val_gbm$Observed)
nRMSEV_gbm <- RMSEV_gbm * 100 / (max(data_val_gbm$Observed) - min(data_val_gbm$Observed))
MAEV_gbm <- MAE(pred = data_val_gbm$Predicted, obs = data_val_gbm$Observed)


# Efficiency (EF) calculation for training and test sets
EF_cal_gbm <- 1 - sum((data_cal_gbm$Observed - data_cal_gbm$Predicted)^2) / sum((data_cal_gbm$Observed - mean(data_cal_gbm$Observed))^2)
EF_val_gbm <- 1 - sum((data_val_gbm$Observed - data_val_gbm$Predicted)^2) / sum((data_val_gbm$Observed - mean(data_val_gbm$Observed))^2)

# Combine all metrics into one dataframe
metrics_df_gbm <- data.frame(
  Set = c("Training", "Test"),
  R_squared = c(r_squared_cal_gbm, r_squared_val_gbm),
  RMSE = c(RMSEC_gbm, RMSEV_gbm),
  nRMSE = c(nRMSEC_gbm, nRMSEV_gbm),
  MAE = c(MAEC_gbm, MAEV_gbm),
  EF = c(EF_cal_gbm, EF_val_gbm)
)

print(metrics_df_gbm)

# Save metrics to CSV
write.csv(metrics_df_gbm, file = "Wheat_GBM_SB_model_metrics_Rust.csv", row.names = FALSE)

# Combine actual and predicted values
data_cal_gbm <- data.frame(Predicted = gbm_train_pred, Observed = train_df$DS)
data_val_gbm <- data.frame(Predicted = gbm_test_pred, Observed = test_df$DS)

# Save actual and predicted values to CSV
write.csv(data_cal_gbm, file = "Wheat GBM_calibration_set_predictions_Rust.csv", row.names = FALSE)
write.csv(data_val_gbm, file = "Wheat GBM_validation_set_predictions_Rust.csv", row.names = FALSE)

# Create data frames for ggplot
data_cal_plot_gbm <- data.frame(Observed = data_cal_gbm$Observed, Predicted = data_cal_gbm$Predicted)
data_val_plot_gbm <- data.frame(Observed = data_val_gbm$Observed, Predicted = data_val_gbm$Predicted)

# Calibration Set Plot
calibration_plot_gbm <- ggplot(data_cal_plot_gbm, aes(x = Observed, y = Predicted)) +
  geom_point(color = "blue", size = 3, alpha = 0.7) +  # Adds points to the plot
  geom_smooth(method = "lm", se = FALSE, linetype = "solid", color = "darkblue", size = 2) +  # Adds a solid line for the linear model
  geom_abline(slope = 1, intercept = 0, linetype = "dashed", color = "red") +  # Adds a dashed line for the 1:1 line
  labs(title = "Calibration - Actual vs Predicted (GBM)",  # Adds the plot title
       x = "Observed",  # Adds the x-axis label
       y = "Predicted") +  # Adds the y-axis label
  theme_minimal(base_size = 15) +  # Sets a minimal theme with a base font size
  annotate("text", x = min(data_cal_gbm$Observed) * 0.5,  # Adjust x position for top-left corner
           y = max(data_cal_gbm$Predicted) * 1.2,  # Adjust y position for top-left corner
           label = sprintf("R-squared = %.3f", r_squared_cal_gbm),
           size = 5, color = "black") +  # Annotates the R-squared value in red
  coord_equal(ratio = 1) +  # Sets equal scaling for x and y axes
  theme(plot.title = element_text(hjust = 0.5),  # Centers the plot title
        panel.grid.major = element_blank(),  # Removes major grid lines
        panel.grid.minor = element_blank())  # Removes minor grid lines


# Validation Set Plot
validation_plot_gbm <- ggplot(data_val_plot_gbm, aes(x = Observed, y = Predicted)) +
  geom_point(color = "black", size = 3, alpha = 0.7) +
  geom_smooth(method = "lm", se = FALSE, linetype = "solid", color = "darkgreen", size = 1.5) +
  geom_abline(slope = 1, intercept = 0, linetype = "dashed", color = "red") +
  labs(title = "Validation - Actual vs Predicted (GBM)",
       x = "Observed",
       y = "Predicted") +
  theme_minimal(base_size = 15) +
  annotate("text", x = max(data_val_gbm$Observed) * 0.5, 
           y = max(data_val_gbm$Predicted) * 1.2,
           label = sprintf("R-squared = %.3f", r_squared_val_gbm),
           size = 5, color = "black") +
  coord_equal(ratio = 1) +
  theme(plot.title = element_text(hjust = 0.5),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank())

# Arrange the plots in a 1x2 grid
grid.arrange(calibration_plot_gbm, validation_plot_gbm, ncol = 2)


# Feature Importance
importance <- summary(gbm_model, n.trees = best_trees)

# Print top features contributing to the model
print(importance)



# Extract mean and standard deviation from preProcValues
mean_values <- preProcValues$mean
std_values <- preProcValues$std

# Assuming 'DS' was the column that was standardized
mean_DS <- mean_values["DS"]
std_DS <- std_values["DS"]

# De-standardize the observed and predicted values
de_standardize <- function(x, mean, std) {
  return(x * std + mean)
}

# Apply de-standardization
data_val_gbm$Predicted_de <- de_standardize(data_val_gbm$Predicted, mean_DS, std_DS)
data_val_gbm$Observed_de <- de_standardize(data_val_gbm$Observed, mean_DS, std_DS)

# View the new dataframe with de-standardized values
print(data_val_gbm)

# Optional: Save the new dataframe to a CSV file
write.csv(data_val_gbm, "E:/AMIT/Thesis Work/Data analysis//Wheat Rust gbm_predictions_destandardized.csv", row.names = FALSE)
