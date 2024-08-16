# Set the working directory to a new path
setwd("E:/AMIT/Thesis Work/Data analysis")


# Verify the current working directory
current_dir <- getwd()
print(current_dir)

# Load required packages
#install.packages("glmnet")
library(glmnet)
library(openxlsx)
library(caret)
library(ggplot2)

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

# Train the glmnet model
set.seed(50)
model1 <- train(DS ~ ., data = train_df, preProcess = c("center", "scale"),
                method = "glmnet",
                tuneGrid = expand.grid(alpha = 0, lambda = lambda_vector), 
                trControl = ctrspes, na.action = na.omit)

# Display best tuning parameters
print(model1$bestTune)
print(model1$bestTune$lambda)
print(round(coef(model1$finalModel, model1$bestTune$lambda), 3))
print(coef(model1$finalModel, s = model1$bestTune$lambda))

# Variable importance plot
print(varImp(model1))
plot(varImp(model1, scale = TRUE))

# Predictions
ridge_cal <- predict(model1, train_df)
ridge_val <- predict(model1, test_df)

# Metrics for training set
data_cal <- data.frame(Predicted = ridge_cal, Observed = train_df$DS)
r_squared_cal <- cor(data_cal$Observed, data_cal$Predicted)^2
RMSEC <- RMSE(pred = data_cal$Predicted, obs = data_cal$Observed)
nRMSEC <- RMSEC * 100 /(max(data_cal$Observed) - min(data_cal$Observed))
MAEC <- MAE(pred = data_cal$Predicted, obs = data_cal$Observed)

# Metrics for test set
data_val <- data.frame(Predicted = ridge_val, Observed = test_df$DS)
r_squared_val <- cor(data_val$Observed, data_val$Predicted)^2
RMSEV <- RMSE(pred = data_val$Predicted, obs = data_val$Observed)
nRMSEV <- RMSEV * 100 / (max(data_val$Observed) - min(data_val$Observed))
MAEV <- MAE(pred = data_val$Predicted, obs = data_val$Observed)

# Calculation of percentage error for test set
error <- ((data_val$Observed - data_val$Predicted) / data_val$Observed) * 100

# Efficiency (EF) calculation for training and test sets
EF_cal <- 1 - sum((data_cal$Observed - data_cal$Predicted)^2) / sum((data_cal$Observed - mean(data_cal$Observed))^2)
EF_val <- 1 - sum((data_val$Observed - data_val$Predicted)^2) / sum((data_val$Observed - mean(data_val$Observed))^2)

# Print metrics
print(paste("Training R-squared: ", format(r_squared_cal, digits = 3)))
print(paste("Training RMSE: ", RMSEC))
print(paste("Training nRMSE: ", nRMSEC))
print(paste("Training MAE: ", MAEC))
print(paste("Training EF: ", EF_cal))

print(paste("Test R-squared: ", format(r_squared_val, digits = 3)))
print(paste("Test RMSE: ", RMSEV))
print(paste("Test nRMSE: ", nRMSEV))
print(paste("Test MAE: ", MAEV))
print(paste("Test EF: ", EF_val))

# Combine all metrics into one dataframe
metrics_df <- data.frame(
  Set = c("Training", "Test"),
  R_squared = c(r_squared_cal, r_squared_val),
  RMSE = c(RMSEC, RMSEV),
  nRMSE = c(nRMSEC, nRMSEV),
  MAE = c(MAEC, MAEV),
  EF = c(EF_cal, EF_val)
)

print(metrics_df)
# Save metrics to CSV
write.csv(metrics_df, file = "Wheat Ridge_model_metrics_Rust.csv", row.names = FALSE)

# View the predicted values of test_df
print("Predicted values for test_df:")
print(ridge_val)


# Combine actual and predicted values
data_cal <- data.frame(Predicted = ridge_cal, Observed = train_df$DS)
data_val <- data.frame(Predicted = ridge_val, Observed = test_df$DS)

# Save actual and predicted values to CSV
write.csv(data_cal, file = "Wheat Ridge_calibration_set_predictions_PM.csv", row.names = FALSE)
write.csv(data_val, file = "Wheat Ridge_validation_set_predictions_Rust.csv", row.names = FALSE)

# Print actual and predicted values
print("Calibration Set - Actual vs Predicted")
print(data_cal)
print("Validation Set - Actual vs Predicted")
print(data_val)

###Printing Regression Equation###################################
# Get the best lambda
best_lambda <- model1$bestTune$lambda
# Extract coefficients and sort by absolute values
coefficients <- coef(model1$finalModel, best_lambda)
coefficients <- as.data.frame(as.matrix(coefficients))
coefficients$Feature <- rownames(coefficients)
coefficients <- coefficients[order(-abs(coefficients$s1)), ]

# Select the top 5 coefficients excluding the intercept
top_coefficients <- coefficients[coefficients$Feature != "(Intercept)", ][1:5, ]

# Print the regression equation
intercept <- coefficients[coefficients$Feature == "(Intercept)", "s1"]
equation <- paste("DS =", round(intercept, 3))
for (i in 1:nrow(top_coefficients)) {
  equation <- paste(equation, 
                    ifelse(top_coefficients[i, "s1"] >= 0, "+", "-"),
                    abs(round(top_coefficients[i, "s1"], 3)), "*", top_coefficients[i, "Feature"])
}
cat("Regression Equation (Top 5 Parameters):\n", equation, "\n")

########Plotting Results#######################
library(ggplot2)
library(gridExtra)

# Create data frames for ggplot
data_cal_plot <- data.frame(Observed = data_cal$Observed, Predicted = data_cal$Predicted)
data_val_plot <- data.frame(Observed = data_val$Observed, Predicted = data_val$Predicted)

# Calibration Set Plot
calibration_plot <- ggplot(data_cal_plot, aes(x = Observed, y = Predicted)) +
  geom_point(color = "blue", size = 3, alpha = 0.7) +
  geom_smooth(method = "lm", se = FALSE, linetype = "solid", color = "darkblue", size = 2) +
  geom_abline(slope = 1, intercept = 0, linetype = "dashed", color = "red") +
  labs(title = "Calibration - Actual vs Predicted",
       x = "Observed",
       y = "Predicted") +
  theme_minimal(base_size = 15) +
  annotate("text", x = max(data_cal$Observed) * 0.5, 
           y = max(data_cal$Predicted) * 1.2,
           label = sprintf("R-squared = %.3f", r_squared_cal),
           size = 5, color = "black") +
  coord_equal(ratio = 1) +
  theme(plot.title = element_text(hjust = 0.5),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank())

# Validation Set Plot
validation_plot <- ggplot(data_val_plot, aes(x = Observed, y = Predicted)) +
  geom_point(color = "black", size = 3, alpha = 0.7) +
  geom_smooth(method = "lm", se = FALSE, linetype = "solid", color = "darkgreen", size = 2) +
  geom_abline(slope = 1, intercept = 0, linetype = "dashed", color = "red") +
  labs(title = "Validation - Actual vs Predicted",
       x = "Observed",
       y = "Predicted") +
  theme_minimal(base_size = 15) +
  annotate("text", x = max(data_val$Observed) * 0.5, 
           y = max(data_val$Predicted) * 1.2,
           label = sprintf("R-squared = %.3f", r_squared_val),
           size = 5, color = "black") +
  coord_equal(ratio = 1) +
  theme(plot.title = element_text(hjust = 0.5),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank())

# Arrange the plots in a 1x2 grid
grid.arrange(calibration_plot, validation_plot, ncol = 2)
