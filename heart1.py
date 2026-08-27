import streamlit as st
import pandas as pd
import joblib
model=joblib.load("KNN_heart.pkl")
scaler=joblib.load("scaler.pkl")
expected_columns=joblib.load("columns.pkl")
st.title("Heart prediction by MONU ")
st.markdown("Provide the following details ")
age=st.slider("Age",18,100,40)
sex=st.selectbox("sex",['M','F'])
chest_pain=st.selectbox("ChestpainType",['ATA','NAP','TA','ASY'])
resting_bp=st.number_input("Resting blood pressure (mm hg)",80,200,120)
cholesterol=st.number_input("Cholesterol (mg/dL)",100,600,200)
fasting_bs=st.selectbox("Fasting Blood sugar >120 mg/dl",[0,1])
resting_ecg=st.selectbox("Resting ECG ",["Normal","ST","LVH"])
max_hr=st.slider("max heart Rate",60,220,150)
exercise_angina=st.selectbox("Exercise induced angina ",["Y","N"])
oldpeak=st.slider("oldpeak(ST depression)",0.0,6.0,1.0)
st_slope=st.selectbox("ST slope",["UP","FLAT","DOWN"])
if st.button("predict"):
    raw_input={
        'age':age,
        'resting_bp':resting_bp,
        'Cholesterol':cholesterol,
        'Fasting BS':fasting_bs,
        'Max HR':max_hr,
        'OldPeak':oldpeak,
        'sex_'+ sex:1,
        'chestpainType_'+chest_pain:1,
        'Exercise Angina_'+exercise_angina:1,
        'ST_Slope_'+st_slope:1
}
# Create DataFrame
input_df = pd.DataFrame([raw_input])

# Add missing columns
for col in expected_columns:
    if col not in input_df.columns:
        input_df[col] = 0

# Reorder columns EXACTLY like training
input_df = input_df[expected_columns]

# Scale input
scaled_input = scaler.transform(input_df)

# Predict
prediction = model.predict(scaled_input)[0]
if prediction==1:
    st.error("High Risk of Heart Disease")
else:
    st.success("Low Risk of Heart Disease ")