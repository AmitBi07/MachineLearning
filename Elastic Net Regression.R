setwd("E:/AMIT/Thesis Work/Data analysis")
# Load required packages
library(glmnet)
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

# Lambda vector for tuning
lambda_vector <- 10^seq(5, -5, length = 500)

# Train the glmnet Elastic Net model (alpha = 0.5)
set.seed(50)
model_enet <- train(DS ~ ., data = train_df, preProcess = c("center", "scale"),
                    method = "glmnet",
                    tuneGrid = expand.grid(alpha = 0.5, lambda = lambda_vector), 
                    trControl = ctrspes, na.action = na.omit)

# Display best tuning parameters
print(model_enet$bestTune)
print(model_enet$bestTune$lambda)
print(round(coef(model_enet$finalModel, model_enet$bestTune$lambda), 3))
print(coef(model_enet$finalModel, s = model_enet$bestTune$lambda))

# Variable importance plot
print(varImp(model_enet))
plot(varImp(model_enet, scale = TRUE))

# Predictions
enet_cal <- predict(model_enet, train_df)
enet_val <- predict(model_enet, test_df)

# Metrics for training set
data_cal_enet <- data.frame(Predicted = enet_cal, Observed = train_df$DS)
r_squared_cal_enet <- cor(data_cal_enet$Observed, data_cal_enet$Predicted)^2
RMSEC_enet <- RMSE(pred = data_cal_enet$Predicted, obs = data_cal_enet$Observed)
nRMSEC_enet <- RMSEC_enet * 100 / (max(data_cal_enet$Observed) - min(data_cal_enet$Observed))
MAEC_enet <- MAE(pred = data_cal_enet$Predicted, obs = data_cal_enet$Observed)

# Metrics for test set
data_val_enet <- data.frame(Predicted = enet_val, Observed = test_df$DS)
r_squared_val_enet <- cor(data_val_enet$Observed, data_val_enet$Predicted)^2
RMSEV_enet <- RMSE(pred = data_val_enet$Predicted, obs = data_val_enet$Observed)
nRMSEV_enet <- RMSEV_enet * 100 / (max(data_val_enet$Observed) - min(data_val_enet$Observed))
MAEV_enet <- MAE(pred = data_val_enet$Predicted, obs = data_val_enet$Observed)

# Efficiency (EF) calculation for training and test sets
EF_cal_enet <- 1 - sum((data_cal_enet$Observed - data_cal_enet$Predicted)^2) / sum((data_cal_enet$Observed - mean(data_cal_enet$Observed))^2)
EF_val_enet <- 1 - sum((data_val_enet$Observed - data_val_enet$Predicted)^2) / sum((data_val_enet$Observed - mean(data_val_enet$Observed))^2)

# Combine all metrics into one dataframe
metrics_df_enet <- data.frame(
  Set = c("Training", "Test"),
  R_squared = c(r_squared_cal_enet, r_squared_val_enet),
  RMSE = c(RMSEC_enet, RMSEV_enet),
  nRMSE = c(nRMSEC_enet, nRMSEV_enet),
  MAE = c(MAEC_enet, MAEV_enet),
  EF = c(EF_cal_enet, EF_val_enet)
)

print(metrics_df_enet)
# Save metrics to CSV
write.csv(metrics_df_enet, file = "Wheat ENET_model_metrics_Rust.csv", row.names = FALSE)

# Combine actual and predicted values
data_cal_enet <- data.frame(Predicted = enet_cal, Observed = train_df$DS)
data_val_enet <- data.frame(Predicted = enet_val, Observed = test_df$DS)

# Save actual and predicted values to CSV
write.csv(data_cal_enet, file = "Wheat PM ENET_calibration_set_predictions.csv", row.names = FALSE)
write.csv(data_val_enet, file = "Wheat ENET_validation_set_predictions_Rust.csv", row.names = FALSE)

# Create data frames for ggplot
data_cal_plot_enet <- data.frame(Observed = data_cal_enet$Observed, Predicted = data_cal_enet$Predicted)
data_val_plot_enet <- data.frame(Observed = data_val_enet$Observed, Predicted = data_val_enet$Predicted)

# Calibration Set Plot
calibration_plot_enet <- ggplot(data_cal_plot_enet, aes(x = Observed, y = Predicted)) +
  geom_point(color = "blue", size = 3, alpha = 0.7) +
  geom_smooth(method = "lm", se = FALSE, linetype = "solid", color = "darkblue", size = 2) +
  geom_abline(slope = 1, intercept = 0, linetype = "dashed", color = "red") +
  labs(title = "Calibration - Actual vs Predicted (Elastic Net)",
       x = "Observed",
       y = "Predicted") +
  theme_minimal(base_size = 15) +
  annotate("text", x = max(data_cal_enet$Observed) * 0.5, 
           y = max(data_cal_enet$Predicted) * 1.2,
           label = sprintf("R-squared = %.3f", r_squared_cal_enet),
           size = 5, color = "black") +
  coord_equal(ratio = 1) +
  theme(plot.title = element_text(hjust = 0.5),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank())

# Validation Set Plot
validation_plot_enet <- ggplot(data_val_plot_enet, aes(x = Observed, y = Predicted)) +
  geom_point(color = "black", size = 3, alpha = 0.7) +
  geom_smooth(method = "lm", se = FALSE, linetype = "solid", color = "darkgreen", size = 2) +
  geom_abline(slope = 1, intercept = 0, linetype = "dashed", color = "red") +
  labs(title = "Validation - Actual vs Predicted (Elastic Net)",
       x = "Observed",
       y = "Predicted") +
  theme_minimal(base_size = 15) +
  annotate("text", x = max(data_val_enet$Observed) * 0.5, 
           y = max(data_val_enet$Predicted) * 1.2,
           label = sprintf("R-squared = %.3f", r_squared_val_enet),
           size = 5, color = "black") +
  coord_equal(ratio = 1) +
  theme(plot.title = element_text(hjust = 0.5),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank())

# Arrange the plots in a 1x2 grid
grid.arrange(calibration_plot_enet, validation_plot_enet, ncol = 2)


### Printing Elastic Net Regression Equation ###################################
# Get the best lambda
best_lambda_enet <- model_enet$bestTune$lambda

# Extract coefficients and sort by absolute values
coefficients_enet <- coef(model_enet$finalModel, best_lambda_enet)
coefficients_enet <- as.data.frame(as.matrix(coefficients_enet))
coefficients_enet$Feature <- rownames(coefficients_enet)
coefficients_enet <- coefficients_enet[order(-abs(coefficients_enet$s1)), ]

# Select the top 5 coefficients excluding the intercept
top_coefficients_enet <- coefficients_enet[coefficients_enet$Feature != "(Intercept)", ][1:5, ]

# Print the regression equation
intercept_enet <- coefficients_enet[coefficients_enet$Feature == "(Intercept)", "s1"]
equation_enet <- paste("DS =", round(intercept_enet, 3))
for (i in 1:nrow(top_coefficients_enet)) {
  equation_enet <- paste(equation_enet, 
                         ifelse(top_coefficients_enet[i, "s1"] >= 0, "+", "-"),
                         abs(round(top_coefficients_enet[i, "s1"], 3)), "*", top_coefficients_enet[i, "Feature"])
}
cat("Elastic Net Regression Equation (Top 5 Parameters):\n", equation_enet, "\n")
