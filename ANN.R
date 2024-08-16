# Load required packages
library(caret)
library(nnet)
library(openxlsx)
library(ggplot2)
library(gridExtra)

# Load the data
df <- read.xlsx("C:/Users/hp/Desktop/divyam ann run/timepass.xlsx", sheet = "Result")
df <- as.data.frame(df)
head(df, 2)
str(df)

# Set seed for reproducibility
set.seed(100)

# Create training and test data (70:30 split)
index <- createDataPartition(df$ACTUAL_YIELD, p = 0.7, list = FALSE, times = 1)
train_df <- df[index, ]
test_df <- df[-index, ]

# Remove NA rows from test_df
test_df <- na.omit(test_df)

# TrainControl object for cross-validation
ctrspes <- trainControl(method = "cv", number = 3, savePredictions = "all")

# Train the ANN model
set.seed(50)
model_ann <- train(ACTUAL_YIELD ~ ., data = train_df, 
                   method = "nnet", 
                   trControl = ctrspes, 
                   preProcess = c("center", "scale"),
                   tuneGrid = expand.grid(size = c(1, 3, 5), decay = c(0.1, 0.5, 1)), 
                   linout = TRUE, trace = FALSE)

# Display best tuning parameters
print(model_ann$bestTune)

# Variable importance plot
print(varImp(model_ann))
plot(varImp(model_ann, scale = TRUE),top = 10)

# Predictions
ann_cal <- predict(model_ann, train_df)
ann_val <- predict(model_ann, test_df)

# Metrics for training set
data_cal_ann <- data.frame(Predicted = ann_cal, Observed = train_df$ACTUAL_YIELD)
r_squared_cal_ann <- cor(data_cal_ann$Observed, data_cal_ann$Predicted)^2
RMSEC_ann <- RMSE(pred = data_cal_ann$Predicted, obs = data_cal_ann$Observed)
nRMSEC_ann <- RMSEC_ann * 100 / (max(data_cal_ann$Observed) - min(data_cal_ann$Observed))
MAEC_ann <- MAE(pred = data_cal_ann$Predicted, obs = data_cal_ann$Observed)

# Metrics for test set
data_val_ann <- data.frame(Predicted = ann_val, Observed = test_df$ACTUAL_YIELD)
r_squared_val_ann <- cor(data_val_ann$Observed, data_val_ann$Predicted)^2
RMSEV_ann <- RMSE(pred = data_val_ann$Predicted, obs = data_val_ann$Observed)
nRMSEV_ann <- RMSEV_ann * 100 /(max(data_val_ann$Observed) - min(data_val_ann$Observed))
MAEV_ann <- MAE(pred = data_val_ann$Predicted, obs = data_val_ann$Observed)

# Efficiency (EF) calculation for training and test sets
EF_cal_ann <- 1 - sum((data_cal_ann$Observed - data_cal_ann$Predicted)^2) / sum((data_cal_ann$Observed - mean(data_cal_ann$Observed))^2)
EF_val_ann <- 1 - sum((data_val_ann$Observed - data_val_ann$Predicted)^2) / sum((data_val_ann$Observed - mean(data_val_ann$Observed))^2)

# Combine all metrics into one dataframe
metrics_df_ann <- data.frame(
  Set = c("Training", "Test"),
  R_squared = c(r_squared_cal_ann, r_squared_val_ann),
  RMSE = c(RMSEC_ann, RMSEV_ann),
  nRMSE = c(nRMSEC_ann, nRMSEV_ann),
  MAE = c(MAEC_ann, MAEV_ann),
  EF = c(EF_cal_ann, EF_val_ann)
)

print(metrics_df_ann)
# Save metrics to CSV
write.csv(metrics_df_ann, file = "Rice ANN_model_metrics_Rust.csv", row.names = FALSE)

# Combine actual and predicted values
data_cal_ann <- data.frame(Predicted = ann_cal, Observed = train_df$ACTUAL_YIELD)
data_val_ann <- data.frame(Predicted = ann_val, Observed = test_df$ACTUAL_YIELD)

# Save actual and predicted values to CSV
write.csv(data_cal_ann, file = "Rice ANN_calibration_set_predictions_PM.csv", row.names = FALSE)
write.csv(data_val_ann, file = "Rice ANN_validation_set_predictions_Rust.csv", row.names = FALSE)

# Create data frames for ggplot
data_cal_plot_ann <- data.frame(Observed = data_cal_ann$Observed, Predicted = data_cal_ann$Predicted)
data_val_plot_ann <- data.frame(Observed = data_val_ann$Observed, Predicted = data_val_ann$Predicted)

# Calibration Set Plot
calibration_plot_ann <- ggplot(data_cal_plot_ann, aes(x = Observed, y = Predicted)) +
  geom_point(color = "blue", size = 3, alpha = 0.7) +
  geom_smooth(method = "lm", se = FALSE, linetype = "solid", color = "darkblue", size = 2) +
  geom_abline(slope = 1, intercept = 0, linetype = "dashed", color = "red") +
  labs(title = "Calibration - Actual vs Predicted (ANN)",
       x = "Observed",
       y = "Predicted") +
  theme_minimal(base_size = 15) +
  annotate("text", x = max(data_cal_ann$Observed) * 0.5, 
           y = max(data_cal_ann$Predicted) * 1.0,
           label = sprintf("R2 = %.2f", r_squared_cal_ann),
           size = 5, color = "red") +
  coord_equal(ratio = 1) +
  theme(plot.title = element_text(hjust = 0.5),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank())

# Validation Set Plot
validation_plot_ann <- ggplot(data_val_plot_ann, aes(x = Observed, y = Predicted)) +
  geom_point(color = "black", size = 3, alpha = 0.7) +
  geom_smooth(method = "lm", se = FALSE, linetype = "solid", color = "darkgreen", size = 2) +
  geom_abline(slope = 1, intercept = 0, linetype = "dashed", color = "red") +
  labs(title = "Validation - Actual vs Predicted (ANN)",
       x = "Observed",
       y = "Predicted") +
  theme_minimal(base_size = 15) +
  annotate("text", x = max(data_val_ann$Observed) * 0.5, 
           y = max(data_val_ann$Predicted) * 1.0,
           label = sprintf("R-squared = %.3f", r_squared_val_ann),
           size = 5, color = "red") +
  coord_equal(ratio = 1) +
  theme(plot.title = element_text(hjust = 0.5),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank())

# Arrange the plots in a 1x2 grid
grid.arrange(calibration_plot_ann, validation_plot_ann, ncol = 2)

### Printing ANN Regression Equation ###
# Extract weights and structure
weights_ann <- model_ann$finalModel$wts
structure_ann <- model_ann$finalModel$n

# Print the structure of the ANN
cat("ANN Structure:\n")
cat("Input layer: ", structure_ann[1], "neurons\n")
cat("Hidden layer: ", structure_ann[2], "neurons\n")
cat("Output layer: ", structure_ann[3], "neuron\n")

print(data_val_ann)
# Note: ANN models do not produce a traditional regression equation like linear models.
# However, the structure and weights provide insight into how the model is constructed.


# Load the NeuralNetTools package
library(NeuralNetTools)

# Plot the structure of the ANN model
plotnet(model_ann)



#########Seperate Plot for Calibration and Validation############


# Create data frames for ggplot
data_cal_plot_ann <- data.frame(Observed = data_cal_ann$Observed, Predicted = data_cal_ann$Predicted)
data_val_plot_ann <- data.frame(Observed = data_val_ann$Observed, Predicted = data_val_ann$Predicted)

# Plotting the results of Calibration
plot(data_cal_plot_ann$Observed, data_cal_plot_ann$Predicted, pch = 19, 
     col = rgb(red = 0, green = 0, blue = 0, alpha = 0.5),
     xlab="Observed yield (kg/ha)", 
     ylab="Predicted yield (kg/ha)", 
     main="Calibration")

lm_cal <- lm(data_cal_plot_ann$Predicted ~ data_cal_plot_ann$Observed)
summary(lm_cal)
abline(lm_cal, col="red")

prd1_cal <- predict(lm_cal, newdata = data.frame(x = data_cal_plot_ann$Observed), 
                    interval = "confidence", 
                    level = 0.95, type = "response")

prd2_cal <- predict(lm_cal, newdata = data.frame(x = data_cal_plot_ann$Observed), 
                    interval = "prediction", 
                    level = 0.95, type = "response")

lines(data_cal_plot_ann$Observed, prd1_cal[,2], col="blue", lty=2)
lines(data_cal_plot_ann$Observed, prd1_cal[,3], col="blue", lty=2)
lines(data_cal_plot_ann$Observed, prd2_cal[,2], col="black", lty=1)
lines(data_cal_plot_ann$Observed, prd2_cal[,3], col="black", lty=1)

# Add legend for confidence intervals
legend("bottomright", legend=c("95% Confidence Interval (Mean)", "95% Prediction Interval (Individual)"),
       col=c("blue", "black"), lty=c(2, 1), bty="n")

# Plotting the results of validation
# (code for validation plot here)

# Plotting the results of validation
plot(data_val_plot_ann$Observed, data_val_plot_ann$Predicted, pch = 19,
     col = rgb(red = 0, green = 0, blue = 0, alpha = 0.5),
     xlab="Observed yield (kg/ha)", 
     ylab="Predicted yield (kg/ha)", 
     main="Validation")

lm_val <- lm(data_val_plot_ann$Predicted ~ data_val_plot_ann$Observed)
summary(lm_val)
abline(lm_val, col="red")

prd1_val <- predict(lm_val, newdata = data.frame(x = data_val_plot_ann$Observed), 
                    interval = "confidence", 
                    level = 0.95, type = "response")

prd2_val <- predict(lm_val, newdata = data.frame(x = data_val_plot_ann$Observed), 
                    interval = "prediction", 
                    level = 0.95, type = "response")

lines(data_val_plot_ann$Observed, prd1_val[,2], col="blue", lty=2)
lines(data_val_plot_ann$Observed, prd1_val[,3], col="blue", lty=2)
lines(data_val_plot_ann$Observed, prd2_val[,2], col="black", lty=1)
lines(data_val_plot_ann$Observed, prd2_val[,3], col="black", lty=1)

# Add legend for confidence intervals
# Add legend for lines
legend("bottomright", legend=c("95% Confidence Interval (Mean)", 
                               "95% Prediction Interval (Individual)",
                               "Linear Regression Fit"),
       col=c("blue", "black", "red"), lty=c(2, 1, 1), bty="n")

