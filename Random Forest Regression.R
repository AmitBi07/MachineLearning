# Load required packages
library(randomForest)
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

# Create training and test data (70:30 split)
index <- createDataPartition(df$DS, p = 0.7, list = FALSE, times = 1)
train_df <- df[index, ]
test_df <- df[-index, ]

# Remove NA rows from test_df
test_df <- na.omit(test_df)

# TrainControl object for cross-validation
ctrspes <- trainControl(method = "cv", number = 3, savePredictions = "all")

# Train the Random Forest model
set.seed(50)
model_rf <- train(DS ~ ., data = train_df, 
                  method = "rf", 
                  trControl = ctrspes, 
                  na.action = na.omit)

# Display best tuning parameters
print(model_rf$bestTune)

# Variable importance plot
print(varImp(model_rf))
plot(varImp(model_rf, scale = TRUE))

# Predictions
rf_cal <- predict(model_rf, train_df)
rf_val <- predict(model_rf, test_df)

# Metrics for training set
data_cal_rf <- data.frame(Predicted = rf_cal, Observed = train_df$DS)
r_squared_cal_rf <- cor(data_cal_rf$Observed, data_cal_rf$Predicted)^2
RMSEC_rf <- RMSE(pred = data_cal_rf$Predicted, obs = data_cal_rf$Observed)
nRMSEC_rf <- RMSEC_rf * 100 / (max(data_cal_rf$Observed) - min(data_cal_rf$Observed))
MAEC_rf <- MAE(pred = data_cal_rf$Predicted, obs = data_cal_rf$Observed)

# Metrics for test set
data_val_rf <- data.frame(Predicted = rf_val, Observed = test_df$DS)
r_squared_val_rf <- cor(data_val_rf$Observed, data_val_rf$Predicted)^2
RMSEV_rf <- RMSE(pred = data_val_rf$Predicted, obs = data_val_rf$Observed)
nRMSEV_rf <- RMSEV_rf * 100 / (max(data_val_rf$Observed) - min(data_val_rf$Observed))
MAEV_rf <- MAE(pred = data_val_rf$Predicted, obs = data_val_rf$Observed)

# Efficiency (EF) calculation for training and test sets
EF_cal_rf <- 1 - sum((data_cal_rf$Observed - data_cal_rf$Predicted)^2) / sum((data_cal_rf$Observed - mean(data_cal_rf$Observed))^2)
EF_val_rf <- 1 - sum((data_val_rf$Observed - data_val_rf$Predicted)^2) / sum((data_val_rf$Observed - mean(data_val_rf$Observed))^2)

# Combine all metrics into one dataframe
metrics_df_rf <- data.frame(
  Set = c("Training", "Test"),
  R_squared = c(r_squared_cal_rf, r_squared_val_rf),
  RMSE = c(RMSEC_rf, RMSEV_rf),
  nRMSE = c(nRMSEC_rf, nRMSEV_rf),
  MAE = c(MAEC_rf, MAEV_rf),
  EF = c(EF_cal_rf, EF_val_rf)
)

print(metrics_df_rf)

# Save metrics to CSV
write.csv(metrics_df_rf, file = "Wheat RF_model_metrics_Rust.csv", row.names = FALSE)

# Combine actual and predicted values
data_cal_rf <- data.frame(Predicted = rf_cal, Observed = train_df$DS)
data_val_rf <- data.frame(Predicted = rf_val, Observed = test_df$DS)

# Save actual and predicted values to CSV
write.csv(data_cal_rf, file = "RF_calibration_set_predictions_SB.csv", row.names = FALSE)
write.csv(data_val_rf, file = "Wheat RF_validation_set_predictions_Rust.csv", row.names = FALSE)

# Create data frames for ggplot
data_cal_plot_rf <- data.frame(Observed = data_cal_rf$Observed, Predicted = data_cal_rf$Predicted)
data_val_plot_rf <- data.frame(Observed = data_val_rf$Observed, Predicted = data_val_rf$Predicted)

# Calibration Set Plot
calibration_plot_rf <- ggplot(data_cal_plot_rf, aes(x = Observed, y = Predicted)) +
  geom_point(color = "blue", size = 3, alpha = 0.7) +
  geom_smooth(method = "lm", se = FALSE, linetype = "solid", color = "darkblue", size = 1) +
  geom_abline(slope = 1, intercept = 0, linetype = "dashed", color = "red") +
  labs(title = "Calibration - Actual vs Predicted (Random Forest)",
       x = "Observed",
       y = "Predicted") +
  theme_minimal(base_size = 15) +
  annotate("text", x = max(data_cal_rf$Observed) * 0.5, 
           y = max(data_cal_rf$Predicted) * 1.2,
           label = sprintf("R-squared = %.3f", r_squared_cal_rf),
           size = 5, color = "red") +
  coord_equal(ratio = 1) +
  theme(plot.title = element_text(hjust = 0.5),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank())

# Validation Set Plot
validation_plot_rf <- ggplot(data_val_plot_rf, aes(x = Observed, y = Predicted)) +
  geom_point(color = "green", size = 3, alpha = 0.7) +
  geom_smooth(method = "lm", se = FALSE, linetype = "solid", color = "darkgreen", size = 1) +
  geom_abline(slope = 1, intercept = 0, linetype = "dashed", color = "red") +
  labs(title = "Validation - Actual vs Predicted (Random Forest)",
       x = "Observed",
       y = "Predicted") +
  theme_minimal(base_size = 15) +
  annotate("text", x = max(data_val_rf$Observed) * 0.5, 
           y = max(data_val_rf$Predicted) * 1.2,
           label = sprintf("R-squared = %.3f", r_squared_val_rf),
           size = 5, color = "red") +
  coord_equal(ratio = 1) +
  theme(plot.title = element_text(hjust = 0.5),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank())

# Arrange the plots in a 1x2 grid
grid.arrange(calibration_plot_rf, validation_plot_rf, ncol = 2)

### Printing Random Forest Regression Equation ###
# Extract important variables and their importance
importance_rf <- varImp(model_rf)
importance_rf <- as.data.frame(importance_rf$importance)
importance_rf$Feature <- rownames(importance_rf)
importance_rf <- importance_rf[order(-importance_rf$Overall), ]

# Select the top 5 important variables
top_importance_rf <- importance_rf[1:5, ]

# Print the importance of the top 5 variables
print("Top 5 important variables:")
print(top_importance_rf)

