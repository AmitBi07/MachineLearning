# Set the working directory to a new path
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

# Train the glmnet Lasso model (alpha = 1)
set.seed(50)
model_lasso <- train(DS ~ ., data = train_df, preProcess = c("center", "scale"),
                     method = "glmnet",
                     tuneGrid = expand.grid(alpha = 1, lambda = lambda_vector), 
                     trControl = ctrspes, na.action = na.omit)

# Display best tuning parameters
print(model_lasso$bestTune)
print(model_lasso$bestTune$lambda)
print(round(coef(model_lasso$finalModel, model_lasso$bestTune$lambda), 3))
print(coef(model_lasso$finalModel, s = model_lasso$bestTune$lambda))

# Variable importance plot
print(varImp(model_lasso))
plot(varImp(model_lasso, scale = TRUE))

# Predictions
lasso_cal <- predict(model_lasso, train_df)
lasso_val <- predict(model_lasso, test_df)

# Metrics for training set
data_cal_lasso <- data.frame(Predicted = lasso_cal, Observed = train_df$DS)
r_squared_cal_lasso <- cor(data_cal_lasso$Observed, data_cal_lasso$Predicted)^2
RMSEC_lasso <- RMSE(pred = data_cal_lasso$Predicted, obs = data_cal_lasso$Observed)
nRMSEC_lasso <- RMSEC_lasso * 100 / (max(data_cal_lasso$Observed) - min(data_cal_lasso$Observed))
MAEC_lasso <- MAE(pred = data_cal_lasso$Predicted, obs = data_cal_lasso$Observed)

# Metrics for test set
data_val_lasso <- data.frame(Predicted = lasso_val, Observed = test_df$DS)
r_squared_val_lasso <- cor(data_val_lasso$Observed, data_val_lasso$Predicted)^2
RMSEV_lasso <- RMSE(pred = data_val_lasso$Predicted, obs = data_val_lasso$Observed)
nRMSEV_lasso <- RMSEV_lasso * 100 / (max(data_val_lasso$Observed) - min(data_val_lasso$Observed))
MAEV_lasso <- MAE(pred = data_val_lasso$Predicted, obs = data_val_lasso$Observed)

# Efficiency (EF) calculation for training and test sets
EF_cal_lasso <- 1 - sum((data_cal_lasso$Observed - data_cal_lasso$Predicted)^2) / sum((data_cal_lasso$Observed - mean(data_cal_lasso$Observed))^2)
EF_val_lasso <- 1 - sum((data_val_lasso$Observed - data_val_lasso$Predicted)^2) / sum((data_val_lasso$Observed - mean(data_val_lasso$Observed))^2)


# Combine all metrics into one dataframe
metrics_df_lasso <- data.frame(
  Set = c("Training", "Test"),
  R_squared = c(r_squared_cal_lasso, r_squared_val_lasso),
  RMSE = c(RMSEC_lasso, RMSEV_lasso),
  nRMSE = c(nRMSEC_lasso, nRMSEV_lasso),
  MAE = c(MAEC_lasso, MAEV_lasso),
  EF = c(EF_cal_lasso, EF_val_lasso)
)

print(metrics_df_lasso)

# Save metrics to CSV
write.csv(metrics_df_lasso, file = "Wheat Lasso_model_metrics_Rust.csv", row.names = FALSE)

# Combine actual and predicted values
data_cal_lasso <- data.frame(Predicted = lasso_cal, Observed = train_df$DS)
data_val_lasso <- data.frame(Predicted = lasso_val, Observed = test_df$DS)

# Save actual and predicted values to CSV
write.csv(data_cal_lasso, file = "Wheat Lasso_calibration_set_predictions_PM.csv", row.names = FALSE)
write.csv(data_val_lasso, file = "Wheat Lasso_validation_set_predictions_Rust.csv", row.names = FALSE)

# Create data frames for ggplot
data_cal_plot_lasso <- data.frame(Observed = data_cal_lasso$Observed, Predicted = data_cal_lasso$Predicted)
data_val_plot_lasso <- data.frame(Observed = data_val_lasso$Observed, Predicted = data_val_lasso$Predicted)

# Calibration Set Plot
calibration_plot_lasso <- ggplot(data_cal_plot_lasso, aes(x = Observed, y = Predicted)) +
  geom_point(color = "blue", size = 3, alpha = 0.7) +
  geom_smooth(method = "lm", se = FALSE, linetype = "solid", color = "darkblue", size = 2) +
  geom_abline(slope = 1, intercept = 0, linetype = "dashed", color = "red") +
  labs(title = "Calibration - Actual vs Predicted (Lasso)",
       x = "Observed",
       y = "Predicted") +
  theme_minimal(base_size = 15) +
  annotate("text", x = max(data_cal_lasso$Observed) * 0.5, 
           y = max(data_cal_lasso$Predicted) * 1.2,
           label = sprintf("R-squared = %.3f", r_squared_cal_lasso),
           size = 5, color = "black") +
  coord_equal(ratio = 1) +
  theme(plot.title = element_text(hjust = 0.5),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank())

# Validation Set Plot
validation_plot_lasso <- ggplot(data_val_plot_lasso, aes(x = Observed, y = Predicted)) +
  geom_point(color = "black", size = 3, alpha = 0.7) +
  geom_smooth(method = "lm", se = FALSE, linetype = "solid", color = "darkgreen", size = 2) +
  geom_abline(slope = 1, intercept = 0, linetype = "dashed", color = "red") +
  labs(title = "Validation - Actual vs Predicted (Lasso)",
       x = "Observed",
       y = "Predicted") +
  theme_minimal(base_size = 15) +
  annotate("text", x = max(data_val_lasso$Observed) * 0.5, 
           y = max(data_val_lasso$Predicted) * 1.2,
           label = sprintf("R-squared = %.3f", r_squared_val_lasso),
           size = 5, color = "black") +
  coord_equal(ratio = 1) +
  theme(plot.title = element_text(hjust = 0.5),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank())

# Arrange the plots in a 1x2 grid
grid.arrange(calibration_plot_lasso, validation_plot_lasso, ncol = 2)


### Printing Lasso Regression Equation ###################################
# Get the best lambda
best_lambda_lasso <- model_lasso$bestTune$lambda

# Extract coefficients and sort by absolute values
coefficients_lasso <- coef(model_lasso$finalModel, best_lambda_lasso)
coefficients_lasso <- as.data.frame(as.matrix(coefficients_lasso))
coefficients_lasso$Feature <- rownames(coefficients_lasso)
coefficients_lasso <- coefficients_lasso[order(-abs(coefficients_lasso$s1)), ]

# Select the top 5 coefficients excluding the intercept
top_coefficients_lasso <- coefficients_lasso[coefficients_lasso$Feature != "(Intercept)", ][1:5, ]

# Print the regression equation
intercept_lasso <- coefficients_lasso[coefficients_lasso$Feature == "(Intercept)", "s1"]
equation_lasso <- paste("DS =", round(intercept_lasso, 3))
for (i in 1:nrow(top_coefficients_lasso)) {
  equation_lasso <- paste(equation_lasso, 
                          ifelse(top_coefficients_lasso[i, "s1"] >= 0, "+", "-"),
                          abs(round(top_coefficients_lasso[i, "s1"], 3)), "*", top_coefficients_lasso[i, "Feature"])
}
cat("Lasso Regression Equation (Top 5 Parameters):\n", equation_lasso, "\n")

