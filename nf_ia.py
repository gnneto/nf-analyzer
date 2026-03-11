import pandas as pd
from sklearn.ensemble import IsolationForest

df = pd.read_excel("notas.xlsx")

X = df[["Valor total NF"]]

model = IsolationForest(contamination=0.05)
df["anomalia"] = model.fit_predict(X)

suspeitas = df[df["anomalia"] == -1]

print(suspeitas)