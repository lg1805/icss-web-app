import joblib

MODEL_PATH = "random_forest_model.pkl"

try:
    model = joblib.load(MODEL_PATH)
    print("Model loaded successfully!")
    print(type(model))
except Exception as e:
    print(f"Error loading model: {e}")
